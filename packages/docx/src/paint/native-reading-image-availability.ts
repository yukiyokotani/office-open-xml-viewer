import { imageNaturalSize } from '@silurus/ooxml-core';
import type { LayoutRect, NativeReadingImagePlan } from '../layout/types.js';
import { isOwnedPaintResourceSession, isUnavailablePaintResourceHandle, type PaintResourceSession, type ResolvedPaintResource } from './resource-session.js';

/** Identity/admission only. Source sealing and geometry are acquired and
 * validated by layout; paint cannot re-import that authority or dry-render a
 * bounds proof. Match every descriptor field that can change decode/draw,
 * including forbidden substitutions, before resolving a visible reading page. */
export function assertNativeReadingImageBinding(
  resource: ResolvedPaintResource<'image' | 'picture-bullet'>, plan: NativeReadingImagePlan, rect: LayoutRect,
): void {
  const d = resource.descriptor, s = plan.source;
  if (d.kind !== 'image' || d.resourceKey !== s.resourceKey || d.partPath !== s.partPath || d.mimeType !== s.mimeType
    || (d.documentOrder ?? null) !== s.documentOrder
    || d.intrinsicSize.widthPt !== s.intrinsicSize.widthPt || d.intrinsicSize.heightPt !== s.intrinsicSize.heightPt
    || d.svgImagePath !== undefined || d.colorReplaceFrom !== undefined || d.duotone !== undefined
    || (d.rotation ?? null) !== s.rotation || (d.flipH ?? null) !== s.flipH
    || (d.flipV ?? null) !== s.flipV || (d.alpha ?? null) !== s.alpha
    || (d.srcRect === undefined) !== (s.srcRect === null)) {
    throw new Error('Native reading image descriptor binding mismatch');
  }
  if (d.srcRect && s.srcRect) for (const edge of ['l', 't', 'r', 'b'] as const) {
    if (d.srcRect[edge] !== s.srcRect[edge]) throw new Error('Native reading image crop binding mismatch');
  }
  if (!Object.values(rect).every(Number.isFinite) || rect.widthPt !== plan.projection.width || rect.heightPt !== plan.projection.height) {
    throw new Error('Native reading image command dimensions mismatch');
  }
}

/** After decoded resource acquisition, before target clear. This returns no
 * geometry and invokes no painter; the narrow class admits a live same-realm
 * immutable bitmap, including when its retained crop/alpha paints no ink. */
export function assertNativeReadingImageAvailable(
  session: PaintResourceSession, plan: NativeReadingImagePlan, rect: LayoutRect,
): void {
  if (!isOwnedPaintResourceSession(session)) throw new Error('Native reading image requires a realm-owned resource session');
  const resource = session.resolve(plan.source.resourceKey, 'image');
  assertNativeReadingImageBinding(resource, plan, rect);
  const image = resource.handle;
  if (isUnavailablePaintResourceHandle(image) || typeof ImageBitmap === 'undefined' || !(image instanceof ImageBitmap)) {
    throw new Error('Native reading image requires an available immutable decoded ImageBitmap');
  }
  const { w, h } = imageNaturalSize(image);
  if (!Number.isSafeInteger(w) || !Number.isSafeInteger(h) || w <= 0 || h <= 0) throw new Error('Native reading image drawable has invalid or closed dimensions');
  const crop = plan.projection.crop;
  if (crop && crop.sourceWidth > 0 && crop.sourceHeight > 0 && crop.dwFraction > 0 && crop.dhFraction > 0
    && ![crop.sourceX * w, crop.sourceY * h, crop.sourceWidth * w, crop.sourceHeight * h].every(Number.isFinite)) {
    throw new Error('Native reading image decoded grid overflow');
  }
}
