import { imagePaintProjection } from '@silurus/ooxml-core';
import type { ImagePaintProjection } from '@silurus/ooxml-core';
import type { ImagePaintResourceDescriptor, LayoutRect, NativeReadingImagePlan } from './types.js';
import { isDeepFrozenPlainDataRoot, snapshotPlainData } from './plain-data.js';
import { composeAffine } from './affine.js';
import { transformRect } from './coordinate-space.js';

/** Source acquisition owns the authoritative plain-data seal. Paint consumes
 * this cloneable projection, never imports the layout brand or collects bounds.
 * Strict intermediate-overflow checks live in the same shared crop/transform
 * owner used by canonical ordinary painting. This is a full-frame library
 * reading allocation, not an Office automatic exclusion contour. */
export function acquireNativeReadingImagePlan(
  descriptor: ImagePaintResourceDescriptor, width: number, height: number,
): NativeReadingImagePlan {
  if (!isDeepFrozenPlainDataRoot(descriptor)) throw new Error('Native image plan requires an internally sealed descriptor');
  if (descriptor.kind !== 'image' || descriptor.svgImagePath !== undefined
    || descriptor.colorReplaceFrom !== undefined || descriptor.duotone !== undefined) {
    throw new Error('Native image substitution/color effects require their decoded-owner proof');
  }
  const source = {
    resourceKey: descriptor.resourceKey, partPath: descriptor.partPath, mimeType: descriptor.mimeType,
    documentOrder: descriptor.documentOrder ?? null,
    intrinsicSize: { ...descriptor.intrinsicSize }, srcRect: descriptor.srcRect ? { ...descriptor.srcRect } : null,
    rotation: descriptor.rotation ?? null, flipH: descriptor.flipH ?? null,
    flipV: descriptor.flipV ?? null, alpha: descriptor.alpha ?? null,
  };
  const projection = imagePaintProjection(width, height, source.srcRect, descriptor);
  return snapshotPlainData({ source, projection }, 'native reading image plan');
}

function sameProjection(actual: ImagePaintProjection, expected: ImagePaintProjection): boolean {
  for (const field of ['width', 'height', 'transformed', 'centerX', 'centerY', 'rotationRad', 'scaleX', 'scaleY', 'alpha'] as const) {
    if (actual[field] !== expected[field]) return false;
  }
  for (const field of ['x', 'y', 'width', 'height'] as const) if (actual.destination[field] !== expected.destination[field]) return false;
  if ((actual.crop === null) !== (expected.crop === null)) return false;
  if (actual.crop && expected.crop) for (const field of ['sourceX', 'sourceY', 'sourceWidth', 'sourceHeight', 'dxFraction', 'dyFraction', 'dwFraction', 'dhFraction'] as const) {
    if (actual.crop[field] !== expected.crop[field]) return false;
  }
  return true;
}

/** Finished-layout validation uses the same constructor, including after
 * structured clone/resealing. It does not assume descendant WeakSet branding
 * survived transport. Local command dimensions/plan/source must still agree. */
export function assertNativeReadingImagePlan(plan: NativeReadingImagePlan, rect: LayoutRect): void {
  const { source, projection } = plan;
  if (![source.resourceKey, source.partPath, source.mimeType].every(value => typeof value === 'string' && value.length > 0)
    || !Object.values(rect).every(Number.isFinite) || rect.widthPt < 0 || rect.heightPt < 0
    || !Object.values(source.intrinsicSize).every(value => Number.isFinite(value) && value >= 0)
    || (source.documentOrder !== null && (!Number.isSafeInteger(source.documentOrder) || source.documentOrder < 0))) {
    throw new Error('Native image retained source/destination mismatch');
  }
  const transform = { ...(source.rotation === null ? {} : { rotation: source.rotation }),
    ...(source.flipH === null ? {} : { flipH: source.flipH }), ...(source.flipV === null ? {} : { flipV: source.flipV }),
    ...(source.alpha === null ? {} : { alpha: source.alpha }) };
  const expected = imagePaintProjection(rect.widthPt, rect.heightPt, source.srcRect, transform);
  if (!sameProjection(projection, expected)) throw new Error('Native image retained projection mismatch');
}

/** Layout alone computes the enclosure of the canonical full destination.
 * Empty crop/alpha retains allocation; live drawable admission happens later,
 * before clear, without invoking the painter or deriving geometry. */
export function nativeReadingImageFrame(plan: NativeReadingImagePlan, rect: LayoutRect): LayoutRect {
  assertNativeReadingImagePlan(plan, rect);
  const p = plan.projection;
  if (!p.transformed) return { ...rect };
  const c = Math.cos(p.rotationRad), s = Math.sin(p.rotationRad);
  const transform = composeAffine({ a: 1, b: 0, c: 0, d: 1, e: rect.xPt + p.centerX, f: rect.yPt + p.centerY },
    composeAffine({ a: c, b: s, c: -s, d: c, e: 0, f: 0 }, { a: p.scaleX, b: 0, c: 0, d: p.scaleY, e: 0, f: 0 }));
  const result = transformRect(transform, { xPt: p.destination.x, yPt: p.destination.y, widthPt: p.width, heightPt: p.height });
  if (!Object.values(result).every(Number.isFinite)) throw new Error('Native image allocation overflow');
  return result;
}
