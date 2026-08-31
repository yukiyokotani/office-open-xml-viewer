import { describe, it, expect } from 'vitest';
import { zoomStepScale, anchoredZoomOffset } from './zoom.js';

/**
 * Ctrl/⌘ + wheel (and trackpad pinch, which the browser reports as a ctrl-wheel)
 * zoom. The old handler ignored `deltaY` magnitude and added a fixed ±0.1 per
 * event, so a trackpad pinch — which fires a high-frequency stream of small
 * wheel events — zoomed far too fast. `zoomStepScale` makes the step
 * exponential in a mode-normalized wheel distance, so fine-grained trackpad
 * input follows gesture distance while oversized discrete wheel events are
 * bounded. One conventional wheel step is deliberately gentle (10%) so a mouse
 * wheel remains adjustable instead of jumping across most of the zoom range.
 */
describe('zoomStepScale (ctrl/pinch zoom)', () => {
  it('maps one conventional pixel-mode wheel step to a 10% scale change', () => {
    expect(zoomStepScale(1, -100, 0)).toBeCloseTo(1.1, 10);
    expect(zoomStepScale(1, 100, 0)).toBeCloseTo(1 / 1.1, 10);
  });

  it('normalizes line- and page-mode wheel deltas to the same step', () => {
    expect(zoomStepScale(1, -3, 1)).toBeCloseTo(1.1, 10);
    expect(zoomStepScale(1, -1, 2)).toBeCloseTo(1.1, 10);
  });

  it('defaults to pixel mode for callers that omit deltaMode', () => {
    expect(zoomStepScale(1, -100)).toBeCloseTo(1.1, 10);
  });

  it('caps an oversized discrete wheel event to one step', () => {
    expect(zoomStepScale(1, -1000, 0)).toBeCloseTo(1.1, 10);
    expect(zoomStepScale(1, 1000, 0)).toBeCloseTo(1 / 1.1, 10);
  });

  it('scrolling up / pinching out (deltaY < 0) zooms in', () => {
    expect(zoomStepScale(1, -10)).toBeGreaterThan(1);
  });

  it('scrolling down / pinching in (deltaY > 0) zooms out', () => {
    expect(zoomStepScale(1, 10)).toBeLessThan(1);
  });

  it('honors deltaY magnitude (a bigger delta zooms more)', () => {
    const small = zoomStepScale(1, -2) - 1;
    const big = zoomStepScale(1, -20) - 1;
    expect(big).toBeGreaterThan(small);
  });

  it('is resolution-independent: two small events ≈ one event of their sum', () => {
    const twoSteps = zoomStepScale(zoomStepScale(1, -5), -5);
    const oneStep = zoomStepScale(1, -10);
    expect(twoSteps).toBeCloseTo(oneStep, 10);
  });

  it('is symmetric: zooming in then out by the same delta returns to start', () => {
    expect(zoomStepScale(zoomStepScale(1, -8), 8)).toBeCloseTo(1, 10);
  });

  it('scales relative to the current zoom (multiplicative, not additive)', () => {
    // Same delta from 200% must move proportionally more than from 100%.
    const from1 = zoomStepScale(1, -10) - 1;
    const from2 = zoomStepScale(2, -10) - 2;
    expect(from2).toBeCloseTo(from1 * 2, 10);
  });
});

/**
 * anchoredZoomOffset — the pointer-anchored ("zoom toward the cursor") scroll
 * correction. The invariant it exists to guarantee: the content pixel under the
 * gesture anchor before the zoom is under the SAME anchor after the zoom. Screen
 * position of that content pixel is `anchor + (c − scrollNew)` where `c` is its
 * scaled-content coordinate; the assertions below verify it lands back on `anchor`.
 */
describe('anchoredZoomOffset (pointer-anchored zoom)', () => {
  const NO_CLAMP = { maxScroll: Number.POSITIVE_INFINITY };

  /** Screen offset of the content pixel that was under `anchor` at the OLD scale,
   *  measured after applying `scrollNew` at the NEW scale. Should equal `anchor`. */
  const screenAfter = (
    scrollOld: number,
    anchor: number,
    scaleOld: number,
    scaleNew: number,
    scrollNew: number,
  ) => {
    const c = (scrollOld + anchor) * (scaleNew / scaleOld); // pinned content px @ new scale
    return c - scrollNew;
  };

  it('keeps the point under the anchor fixed when zooming IN', () => {
    const scrollOld = 300;
    const anchor = 200;
    const scrollNew = anchoredZoomOffset(scrollOld, anchor, 1, 2, NO_CLAMP);
    expect(screenAfter(scrollOld, anchor, 1, 2, scrollNew)).toBeCloseTo(anchor, 9);
  });

  it('keeps the point under the anchor fixed when zooming OUT', () => {
    const scrollOld = 800;
    const anchor = 150;
    const scrollNew = anchoredZoomOffset(scrollOld, anchor, 2, 1, NO_CLAMP);
    expect(screenAfter(scrollOld, anchor, 2, 1, scrollNew)).toBeCloseTo(anchor, 9);
  });

  it('is the identity when the scale does not change (no drift on a no-op zoom)', () => {
    expect(anchoredZoomOffset(432, 210, 1.3, 1.3, NO_CLAMP)).toBeCloseTo(432, 9);
  });

  it('round-trips: zoom in by r then out by 1/r returns the original offset', () => {
    const scrollOld = 500;
    const anchor = 120;
    const zoomedIn = anchoredZoomOffset(scrollOld, anchor, 1, 1.6, NO_CLAMP);
    const back = anchoredZoomOffset(zoomedIn, anchor, 1.6, 1, NO_CLAMP);
    expect(back).toBeCloseTo(scrollOld, 9);
  });

  it('anchor at the region start (0) reduces to an origin-anchored rescale', () => {
    // scrollNew = scrollOld · ratio when anchor = 0.
    expect(anchoredZoomOffset(400, 0, 1, 2, NO_CLAMP)).toBeCloseTo(800, 9);
    expect(anchoredZoomOffset(400, 0, 2, 1, NO_CLAMP)).toBeCloseTo(200, 9);
  });

  it('clamps to 0 (never scrolls above the content top)', () => {
    // Zooming out near the top would want a negative offset; pin to 0.
    const scrollNew = anchoredZoomOffset(10, 200, 2, 1, NO_CLAMP);
    expect(scrollNew).toBe(0);
  });

  it('clamps to maxScroll (never scrolls past the content bottom)', () => {
    // A large zoom-in near the bottom would overshoot the scrollable extent.
    const scrollNew = anchoredZoomOffset(900, 400, 1, 4, { maxScroll: 1000 });
    expect(scrollNew).toBe(1000);
  });

  it('treats a degenerate negative maxScroll (content < viewport) as 0', () => {
    expect(anchoredZoomOffset(0, 100, 1, 2, { maxScroll: -50 })).toBe(0);
  });

  it('does not divide by zero when the old scale is 0 (returns a clamped no-op)', () => {
    expect(anchoredZoomOffset(120, 60, 0, 2, NO_CLAMP)).toBe(120);
  });
});
