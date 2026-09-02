import { drawImageCropped, paintOptionalImagePlaceholder, renderChart } from '@silurus/ooxml-core';
import type {
  ChartImageLookup,
  ChartThreeDRenderer,
  ChartRegionMapRenderer,
  ChartExRenderer,
} from '@silurus/ooxml-core';
import type {
  ImagePaintResourceDescriptor,
  LayoutRect,
} from '../layout/types.js';
import {
  isUnavailablePaintResourceHandle,
  type ResolvedPaintResource,
} from './resource-session.js';
import type {
  CanvasPaintResourceHandlers,
  PaintCanvas2D,
} from './types.js';

type DrawablePaintResource =
  | ResolvedPaintResource<'image'>
  | ResolvedPaintResource<'picture-bullet'>
  | ResolvedPaintResource<'math'>;

function drawableHandle(
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

export function paintImageResource(
  resource: ResolvedPaintResource<'image'> | ResolvedPaintResource<'picture-bullet'>,
  bounds: LayoutRect,
  ctx: PaintCanvas2D,
): void {
  const descriptor = resource.descriptor as ImagePaintResourceDescriptor;
  const image = drawableHandle(resource);
  const placeholder = isUnavailablePaintResourceHandle(resource.handle)
    ? resource.handle.placeholder
    : undefined;
  if (!image && !placeholder) return;
  const draw = (xPt: number, yPt: number): void => {
    if (image) {
      drawImageCropped(
        ctx as CanvasRenderingContext2D,
        image,
        descriptor.srcRect,
        xPt,
        yPt,
        bounds.widthPt,
        bounds.heightPt,
      );
    } else if (placeholder === 'tiff') {
      paintOptionalImagePlaceholder(ctx as CanvasRenderingContext2D, 'tiff', {
        x: xPt,
        y: yPt,
        width: bounds.widthPt,
        height: bounds.heightPt,
      });
    }
  };
  const hasAlpha = descriptor.alpha !== undefined && descriptor.alpha < 1;
  if (hasAlpha) {
    ctx.save();
    ctx.globalAlpha *= descriptor.alpha as number;
  }
  const rotation = descriptor.rotation ?? 0;
  if (rotation === 0 && !descriptor.flipH && !descriptor.flipV) {
    draw(bounds.xPt, bounds.yPt);
  } else {
    ctx.save();
    ctx.translate(
      bounds.xPt + bounds.widthPt / 2,
      bounds.yPt + bounds.heightPt / 2,
    );
    ctx.rotate(rotation * Math.PI / 180);
    ctx.scale(descriptor.flipH ? -1 : 1, descriptor.flipV ? -1 : 1);
    draw(-bounds.widthPt / 2, -bounds.heightPt / 2);
    ctx.restore();
  }
  if (hasAlpha) ctx.restore();
}

function paintDrawableResource(
  resource: ResolvedPaintResource<'math'> | ResolvedPaintResource<'picture-bullet'>,
  bounds: LayoutRect,
  ctx: PaintCanvas2D,
): void {
  const drawable = drawableHandle(resource);
  if (!drawable) {
    if (isUnavailablePaintResourceHandle(resource.handle)
      && resource.handle.placeholder === 'tiff') {
      paintOptionalImagePlaceholder(ctx as CanvasRenderingContext2D, 'tiff', {
        x: bounds.xPt,
        y: bounds.yPt,
        width: bounds.widthPt,
        height: bounds.heightPt,
      });
    }
    return;
  }
  ctx.drawImage(
    drawable,
    bounds.xPt,
    bounds.yPt,
    bounds.widthPt,
    bounds.heightPt,
  );
}

export function createCanonicalCanvasPaintResourceHandlers(
  threeD?: ChartThreeDRenderer,
  regionMap?: ChartRegionMapRenderer,
  chartImageLookup?: ChartImageLookup,
  chartEx?: ChartExRenderer,
): CanvasPaintResourceHandlers {
  return Object.freeze({
  image(resource, bounds, ctx) {
    paintImageResource(resource, bounds, ctx);
  },
  chart(resource, bounds, ctx) {
    // paintLayoutPage has already installed the point-to-device CTM. Passing 1
    // keeps chart font/line point sizes in that same space instead of scaling twice.
    renderChart(
      ctx as CanvasRenderingContext2D,
      resource.descriptor.model as import('@silurus/ooxml-core').ChartModel,
      { x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt },
      1,
      0,
      threeD,
      regionMap,
      chartImageLookup,
      chartEx,
    );
  },
  math(resource, bounds, ctx) {
    paintDrawableResource(resource, bounds, ctx);
  },
  'picture-bullet'(resource, bounds, ctx) {
    paintDrawableResource(resource, bounds, ctx);
  },
  });
}

export const canonicalCanvasPaintResourceHandlers: CanvasPaintResourceHandlers =
  createCanonicalCanvasPaintResourceHandlers();
