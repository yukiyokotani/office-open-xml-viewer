import { drawImageCropped, paintOptionalImagePlaceholder } from '@silurus/ooxml-core';
import type { ImagePaintResourceDescriptor, LayoutRect, NativeReadingImagePlan } from '../layout/types.js';
import { isUnavailablePaintResourceHandle, type ResolvedPaintResource } from './resource-session.js';
import { assertNativeReadingImageBinding } from './native-reading-image-availability.js';
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

/** Thin canonical image adapter. The shared core primitive owns ordinary and
 * retained crop/transform execution; layout supplies a sealed reading plan.
 * Preserve this helper/re-export for image and picture-bullet callers. */
export function paintImageResource(
  resource: ResolvedPaintResource<'image'> | ResolvedPaintResource<'picture-bullet'>,
  bounds: LayoutRect, ctx: PaintCanvas2D, nativeImagePlan?: NativeReadingImagePlan,
): void {
  const descriptor = resource.descriptor as ImagePaintResourceDescriptor;
  if (nativeImagePlan) {
    if (descriptor.kind !== 'image') throw new Error('Reading plan requires an image resource');
    assertNativeReadingImageBinding(resource as ResolvedPaintResource<'image'>, nativeImagePlan, bounds);
  }
  const image = drawableHandle(resource);
  const placeholder = isUnavailablePaintResourceHandle(resource.handle) ? resource.handle.placeholder : undefined;
  if (!image && !placeholder) return;
  const options = nativeImagePlan ? { projection: nativeImagePlan.projection }
    : { transform: { rotation: descriptor.rotation, flipH: descriptor.flipH, flipV: descriptor.flipV, alpha: descriptor.alpha } };
  if (image) {
    drawImageCropped(ctx as CanvasRenderingContext2D, image, descriptor.srcRect,
      bounds.xPt, bounds.yPt, bounds.widthPt, bounds.heightPt, options);
  } else if (placeholder === 'tiff') {
    if (nativeImagePlan) throw new Error('Reading image cannot use a placeholder');
    paintOptionalImagePlaceholder(ctx as CanvasRenderingContext2D, 'tiff', {
      x: bounds.xPt, y: bounds.yPt, width: bounds.widthPt, height: bounds.heightPt,
    }, options);
  }
}
