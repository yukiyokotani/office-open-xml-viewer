import { drawingMLTileOrigin, isDrawingMLTileBrush, isVmlImageBrush } from './vml-image-brush.js';
import { drawImageCropped, paintOptionalImagePlaceholder } from '@silurus/ooxml-core';
import type { ImagePaintResourceDescriptor, LayoutRect } from '../layout/types.js';
import { isUnavailablePaintResourceHandle, type ResolvedPaintResource } from './resource-session.js';
import type { PaintCanvas2D } from './types.js';

type DrawablePaintResource =
  | ResolvedPaintResource<'image'>
  | ResolvedPaintResource<'picture-bullet'>
  | ResolvedPaintResource<'math'>;

export function drawableHandle(
  resource: DrawablePaintResource,
): CanvasImageSource | undefined {
  if (isUnavailablePaintResourceHandle(resource.handle)) return undefined;
  if (resource.handle === undefined || resource.handle === null) {
    throw new Error(
      `Missing ${resource.descriptor.kind} drawable for ${resource.descriptor.resourceKey}`,
    );
  }
  return resource.handle as CanvasImageSource;
}

/** Thin canonical image adapter. The shared core primitive owns crop and
 * transform execution. Preserve this helper/re-export for image and
 * picture-bullet callers. */
export function paintImageResource(
  resource: ResolvedPaintResource<'image'> | ResolvedPaintResource<'picture-bullet'>,
  bounds: LayoutRect, ctx: PaintCanvas2D,
  coverageBounds: LayoutRect = bounds,
): void {
  const descriptor = resource.descriptor as ImagePaintResourceDescriptor;
  if (isDrawingMLTileBrush(resource.handle)) {
    const brush = resource.handle;
    const origin = drawingMLTileOrigin(brush, bounds);
    const sx = brush.tileWidthPt / brush.pixelWidth;
    const sy = brush.tileHeightPt / brush.pixelHeight;
    ctx.save();
    try {
      ctx.globalAlpha *= descriptor.alpha ?? 1;
      const pattern = ctx.createPattern?.(brush.image, 'repeat');
      if (!pattern) throw new Error('Canvas cannot create a retained DrawingML tile pattern');
      ctx.translate(origin.xPt, origin.yPt);
      ctx.scale(sx, sy);
      ctx.fillStyle = pattern;
      ctx.fillRect((coverageBounds.xPt - origin.xPt) / sx, (coverageBounds.yPt - origin.yPt) / sy,
        coverageBounds.widthPt / sx, coverageBounds.heightPt / sy);
    } finally { ctx.restore(); }
    return;
  }
  if (isVmlImageBrush(resource.handle)) {
    const brush = resource.handle;
    const width = brush.width * bounds.widthPt, height = brush.height * bounds.heightPt;
    const x = bounds.xPt + brush.x * bounds.widthPt, y = bounds.yPt + brush.y * bounds.heightPt;
    ctx.save();
    try {
      ctx.globalAlpha *= descriptor.alpha ?? 1;
      if (brush.repeat) {
        const pattern = ctx.createPattern?.(brush.image, 'repeat');
        if (!pattern) throw new Error('Canvas cannot create a retained VML image pattern');
        ctx.translate(x, y);
        ctx.scale(width / brush.pixelWidth, height / brush.pixelHeight);
        ctx.fillStyle = pattern;
        ctx.fillRect((coverageBounds.xPt - x) * brush.pixelWidth / width,
          (coverageBounds.yPt - y) * brush.pixelHeight / height,
          coverageBounds.widthPt * brush.pixelWidth / width, coverageBounds.heightPt * brush.pixelHeight / height);
      } else {
        ctx.drawImage(brush.image, x, y, width, height);
      }
    } finally { ctx.restore(); }
    return;
  }
  const image = drawableHandle(resource);
  const placeholder = isUnavailablePaintResourceHandle(resource.handle) ? resource.handle.placeholder : undefined;
  if (!image && !placeholder) return;
  const options = { transform: { rotation: descriptor.rotation, flipH: descriptor.flipH, flipV: descriptor.flipV, alpha: descriptor.alpha } };
  if (image) {
    drawImageCropped(ctx as CanvasRenderingContext2D, image, descriptor.srcRect,
      bounds.xPt, bounds.yPt, bounds.widthPt, bounds.heightPt, options);
  } else if (placeholder === 'tiff') {
    paintOptionalImagePlaceholder(ctx as CanvasRenderingContext2D, 'tiff', {
      x: bounds.xPt, y: bounds.yPt, width: bounds.widthPt, height: bounds.heightPt,
    }, options);
  }
}
