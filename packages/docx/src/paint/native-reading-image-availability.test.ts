import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import { createPaintResourceRegistry } from '../layout/paint-resources.js';
import { acquireNativeReadingImagePlan } from '../layout/native-reading-image-frame.js';
import type { ImagePaintResourceDescriptor } from '../layout/types.js';
import { createPaintResourceSession, unavailablePaintResourceHandle } from './resource-session.js';
import { assertNativeReadingImageAvailable } from './native-reading-image-availability.js';

// Invented constructor-interface tests only; genuine browser/decode ownership
// remains a separately qualified runtime gate, never proved by these stubs.
class InventedBitmap { width = 100; height = 80; }
const rect = { xPt: 10, yPt: 20, widthPt: 40, heightPt: 20 };
function fixture(extra: Partial<ImagePaintResourceDescriptor> = {}, handle: unknown = new InventedBitmap()) {
  const registry = createPaintResourceRegistry([{ kind: 'image', resourceKey: 'invented-image', partPath: 'invented.png', mimeType: 'image/png', intrinsicSize: { widthPt: 40, heightPt: 20 }, ...extra }]);
  return createPaintResourceSession(registry, [{ resourceKey: 'invented-image', kind: 'image', handle }]);
}
function withConstructor(fn: () => void) {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'ImageBitmap');
  Object.defineProperty(globalThis, 'ImageBitmap', { value: InventedBitmap, configurable: true });
  try { fn(); } finally { if (previous) Object.defineProperty(globalThis, 'ImageBitmap', previous); else Reflect.deleteProperty(globalThis, 'ImageBitmap'); }
}

describe('native reading image resource-only admission', () => {
  it('requires the real session and an available live immutable bitmap even for empty crop', () => withConstructor(() => {
    const owned = fixture({ srcRect: { l: 2, t: 0, r: 0, b: 0 } });
    const plan = acquireNativeReadingImagePlan(owned.resolve('invented-image', 'image').descriptor, 40, 20);
    assert.doesNotThrow(() => assertNativeReadingImageAvailable(owned, plan, rect));
    assert.throws(() => assertNativeReadingImageAvailable({ keys: owned.keys, resolve: owned.resolve }, plan, rect), /realm-owned/);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ srcRect: { l: 2, t: 0, r: 0, b: 0 } }, unavailablePaintResourceHandle('unavailable')), plan, rect), /available immutable/);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ srcRect: { l: 2, t: 0, r: 0, b: 0 } }, { width: 100, height: 80 }), plan, rect), /available immutable/);
    const closed = Object.assign(new InventedBitmap(), { width: 0 });
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ srcRect: { l: 2, t: 0, r: 0, b: 0 } }, closed), plan, rect), /closed dimensions/);
  }));
  it('rejects decoder substitutions, descriptor drift and changed command dimensions', () => withConstructor(() => {
    const owned = fixture(), plan = acquireNativeReadingImagePlan(owned.resolve('invented-image', 'image').descriptor, 40, 20);
    const before = JSON.stringify(plan);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ svgImagePath: 'other.svg' }), plan, rect), /binding mismatch/);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ colorReplaceFrom: '' }), plan, rect), /binding mismatch/);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ duotone: { clr1: '000000', clr2: 'FFFFFF' } }), plan, rect), /binding mismatch/);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ documentOrder: 1 }), plan, rect), /binding mismatch/);
    assert.throws(() => assertNativeReadingImageAvailable(fixture({ intrinsicSize: { widthPt: 10, heightPt: 20 } }), plan, rect), /binding mismatch/);
    assert.throws(() => assertNativeReadingImageAvailable(owned, plan, { ...rect, widthPt: 41 }), /dimensions mismatch/);
    assert.equal(JSON.stringify(plan), before);
  }));
});
