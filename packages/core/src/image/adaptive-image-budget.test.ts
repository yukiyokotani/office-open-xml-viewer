import { describe, expect, it } from 'vitest';
import {
  HARD_MAX_DECODED_IMAGE_BYTES,
  MAX_DECODED_IMAGE_BYTES,
} from './pixel-budget.js';
import {
  normalizeImageResourceOptions,
  planDecodedImageTargets,
} from './adaptive-image-budget.js';

describe('normalizeImageResourceOptions', () => {
  it('defaults to adaptive quality under the standard decoded-byte budget', () => {
    expect(normalizeImageResourceOptions()).toEqual({
      decodedByteBudget: MAX_DECODED_IMAGE_BYTES,
      strategy: 'adaptive',
    });
  });

  it('accepts an application budget up to the internal hard ceiling', () => {
    expect(normalizeImageResourceOptions({
      decodedByteBudget: HARD_MAX_DECODED_IMAGE_BYTES,
      strategy: 'strict',
    })).toEqual({
      decodedByteBudget: HARD_MAX_DECODED_IMAGE_BYTES,
      strategy: 'strict',
    });
  });

  it.each([0, -1, 1.5, Number.NaN, HARD_MAX_DECODED_IMAGE_BYTES + 1])(
    'rejects unsafe decoded-byte budget %s',
    (decodedByteBudget) => {
      expect(() => normalizeImageResourceOptions({ decodedByteBudget }))
        .toThrow(RangeError);
    },
  );
});

describe('planDecodedImageTargets', () => {
  it('reduces display resolution for a default aggregate crossing', () => {
    const policy = normalizeImageResourceOptions();
    const plan = planDecodedImageTargets([
      { key: 'a', targetWidthPx: 4096, targetHeightPx: 4096 },
      { key: 'b', targetWidthPx: 4096, targetHeightPx: 4096 },
      { key: 'c', targetWidthPx: 4096, targetHeightPx: 4096 },
    ], policy);

    expect(plan.degraded).toBe(true);
    expect(plan.plannedBytes).toBeLessThanOrEqual(MAX_DECODED_IMAGE_BYTES);
  });

  it('keeps display-resolution targets unchanged when they fit', () => {
    const plan = planDecodedImageTargets([
      { key: 'a', targetWidthPx: 100, targetHeightPx: 50 },
      { key: 'b', targetWidthPx: 20, targetHeightPx: 10, retainedSurfaceCount: 2 },
    ], { decodedByteBudget: 1_000_000, strategy: 'adaptive' });

    expect(plan.degraded).toBe(false);
    expect(plan.targets.get('a')).toMatchObject({ width: 100, height: 50 });
    expect(plan.targets.get('b')).toMatchObject({ width: 20, height: 10 });
    expect(plan.plannedBytes).toBe((100 * 50 + 20 * 10 * 2) * 4);
  });

  it('preserves established native decoding while the source working set fits', () => {
    const plan = planDecodedImageTargets([{
      key: 'photo',
      targetWidthPx: 400,
      targetHeightPx: 300,
      sourceWidthPx: 4000,
      sourceHeightPx: 3000,
    }], { decodedByteBudget: MAX_DECODED_IMAGE_BYTES, strategy: 'adaptive' });

    expect(plan.targets.get('photo')).toMatchObject({ width: 4000, height: 3000 });
    expect(plan.plannedBytes).toBe(4000 * 3000 * 4);
  });

  it('applies the surface ceiling per surface rather than to their aggregate', () => {
    const plan = planDecodedImageTargets([{
      key: 'effect-photo',
      targetWidthPx: 1000,
      targetHeightPx: 800,
      sourceWidthPx: 5000,
      sourceHeightPx: 4000,
      retainedSurfaceCount: 2,
    }], { decodedByteBudget: 200_000_000, strategy: 'adaptive' });

    expect(plan.targets.get('effect-photo')).toMatchObject({ width: 5000, height: 4000 });
    expect(plan.plannedBytes).toBe(5000 * 4000 * 2 * 4);
  });

  it('switches an oversized but admissible source to its display target', () => {
    const plan = planDecodedImageTargets([{
      key: 'poster',
      targetWidthPx: 1280,
      targetHeightPx: 960,
      sourceWidthPx: 12_090,
      sourceHeightPx: 9_063,
    }], { decodedByteBudget: MAX_DECODED_IMAGE_BYTES, strategy: 'adaptive' });

    expect(plan.targets.get('poster')).toEqual({
      width: 1280,
      height: 960,
      retainedWidth: 1281,
      retainedHeight: 961,
      retainedPixels: 1281 * 961,
    });
    expect(plan.plannedBytes).toBe(1281 * 961 * 4);
  });

  it('allocates one uniform pixels-per-display-pixel scale when the pass is over budget', () => {
    const plan = planDecodedImageTargets([
      { key: 'a', targetWidthPx: 100, targetHeightPx: 100 },
      { key: 'b', targetWidthPx: 100, targetHeightPx: 100 },
    ], { decodedByteBudget: 20_000, strategy: 'adaptive' });

    expect(plan.degraded).toBe(true);
    expect(plan.qualityScale).toBeCloseTo(0.5, 2);
    expect(plan.targets.get('a')).toMatchObject({ width: 50, height: 50 });
    expect(plan.targets.get('b')).toMatchObject({ width: 50, height: 50 });
    expect(plan.plannedBytes).toBeLessThanOrEqual(20_000);
  });

  it('deduplicates repeated cache identities and plans for their largest use', () => {
    const plan = planDecodedImageTargets([
      { key: 'same', targetWidthPx: 100, targetHeightPx: 50 },
      { key: 'same', targetWidthPx: 200, targetHeightPx: 100 },
    ], { decodedByteBudget: 1_000_000, strategy: 'adaptive' });

    expect(plan.targets.size).toBe(1);
    expect(plan.targets.get('same')).toMatchObject({ width: 200, height: 100 });
    expect(plan.plannedBytes).toBe(200 * 100 * 4);
  });

  it('never exceeds the configured byte budget after integer rounding', () => {
    const plan = planDecodedImageTargets(Array.from({ length: 37 }, (_, index) => ({
      key: `image-${index}`,
      targetWidthPx: 333 + index,
      targetHeightPx: 271 + index,
      retainedSurfaceCount: index % 3 === 0 ? 2 : 1,
    })), { decodedByteBudget: 1_234_567, strategy: 'adaptive' });

    expect(plan.degraded).toBe(true);
    expect(plan.plannedBytes).toBeLessThanOrEqual(1_234_567);
  });

  it('rejects when even one pixel per retained surface cannot fit', () => {
    expect(() => planDecodedImageTargets([
      { key: 'a', targetWidthPx: 100, targetHeightPx: 100 },
      { key: 'b', targetWidthPx: 100, targetHeightPx: 100 },
    ], { decodedByteBudget: 4, strategy: 'adaptive' })).toThrow(expect.objectContaining({
      code: 'ooxml-decoded-image-limit',
      metric: 'active-decoded-bytes',
      limit: 4,
      observed: 8,
    }));
  });

  it('rejects strict aggregate crossings before starting decode work', () => {
    expect(() => planDecodedImageTargets([
      { key: 'a', targetWidthPx: 100, targetHeightPx: 100 },
    ], { decodedByteBudget: 100, strategy: 'strict' })).toThrow(expect.objectContaining({
      code: 'ooxml-decoded-image-limit',
      metric: 'active-decoded-bytes',
      limit: 100,
      observed: 40_000,
    }));
  });
});
