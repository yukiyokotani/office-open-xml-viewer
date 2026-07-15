import { describe, expect, it } from 'vitest';
import { canonicalLogicalToPhysical, composeAffine } from '../layout/affine.js';
import { descendPageFrame, pageFrameReentry } from './page-frame.js';
import { translationAffine } from './affine.js';

describe('page-frame adapter', () => {
  it.each([
    { horizontal: 'page', vertical: 'page' },
    { horizontal: 'page', vertical: 'host' },
    { horizontal: 'host', vertical: 'page' },
    { horizontal: 'host', vertical: 'host' },
  ] as const)('re-enters one physical frame after nested translation for $horizontal/$vertical ownership', (axes) => {
    const region = canonicalLogicalToPhysical('vertical-rl', 333);
    const descended = descendPageFrame(
      { currentToPage: region },
      translationAffine(41, 57),
    )!;
    const reentry = pageFrameReentry(descended, {
      coordinateSpace: 'physical-page-points',
      ...axes,
    });
    const final = composeAffine(descended.currentToPage, reentry.currentToTarget);

    expect(final).toEqual({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    expect(final.a * final.d - final.b * final.c).toBe(1);
    expect(Object.values(final).every(Number.isFinite)).toBe(true);
  });
});
