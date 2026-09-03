import { describe, it, expect } from 'vitest';
import {
  ZOOM_STEP_LADDER,
  nextZoomStep,
  prevZoomStep,
  clampScale,
  fitScale,
} from './zoomable.js';

/**
 * IX9 — the shared zoom contract's PURE helpers. The interface itself
 * ({@link import('./zoomable.js').ZoomableViewer}) is a type and is verified by the
 * per-viewer tests; here we lock the ladder stepping, the fit math, and the clamp
 * that every viewer's `zoomIn`/`zoomOut`/`fitWidth`/`fitPage` build on.
 */
describe('ZOOM_STEP_LADDER', () => {
  it('is strictly ascending', () => {
    for (let i = 1; i < ZOOM_STEP_LADDER.length; i++) {
      expect(ZOOM_STEP_LADDER[i]).toBeGreaterThan(ZOOM_STEP_LADDER[i - 1]);
    }
  });

  it('includes 100% and spans 25%–400% (the Excel/Acrobat family)', () => {
    expect(ZOOM_STEP_LADDER).toContain(1);
    expect(ZOOM_STEP_LADDER[0]).toBe(0.25);
    expect(ZOOM_STEP_LADDER[ZOOM_STEP_LADDER.length - 1]).toBe(4);
  });

  it('is frozen (a caller cannot mutate the shared ladder)', () => {
    expect(Object.isFrozen(ZOOM_STEP_LADDER)).toBe(true);
  });
});

describe('nextZoomStep (+ step)', () => {
  it('jumps to the next rung above an on-rung value', () => {
    expect(nextZoomStep(1)).toBe(1.1);
    expect(nextZoomStep(0.25)).toBe(0.33);
    expect(nextZoomStep(2)).toBe(2.5);
  });

  it('snaps an off-ladder value UP to the next rung', () => {
    // 1.03 (a wheel-zoomed value) → the first rung strictly greater, 1.1.
    expect(nextZoomStep(1.03)).toBe(1.1);
    expect(nextZoomStep(0.8)).toBe(0.9);
  });

  it('caps at the top rung', () => {
    expect(nextZoomStep(4)).toBe(4);
    expect(nextZoomStep(10)).toBe(4);
  });

  it('does not re-select a rung it is essentially already on (float slop)', () => {
    expect(nextZoomStep(1 + 1e-9)).toBe(1.1);
  });
});

describe('prevZoomStep (− step)', () => {
  it('jumps to the next rung below an on-rung value', () => {
    expect(prevZoomStep(1)).toBe(0.9);
    expect(prevZoomStep(4)).toBe(3);
    expect(prevZoomStep(0.33)).toBe(0.25);
  });

  it('snaps an off-ladder value DOWN to the next rung', () => {
    expect(prevZoomStep(1.03)).toBe(1);
    expect(prevZoomStep(0.7)).toBe(0.67);
  });

  it('floors at the bottom rung', () => {
    expect(prevZoomStep(0.25)).toBe(0.25);
    expect(prevZoomStep(0.05)).toBe(0.25);
  });

  it('steps below the ladder to an explicit fitted floor', () => {
    expect(prevZoomStep(0.25, 0.05)).toBe(0.05);
    expect(prevZoomStep(0.1, 0.05)).toBe(0.05);
    expect(prevZoomStep(0.05, 0.05)).toBe(0.05);
  });

  it('round-trips +/- symmetrically on-ladder', () => {
    expect(prevZoomStep(nextZoomStep(1))).toBe(1);
    expect(nextZoomStep(prevZoomStep(2))).toBe(2);
  });
});

describe('clampScale', () => {
  it('passes a value inside the range through', () => {
    expect(clampScale(1.5, 0.1, 4)).toBe(1.5);
  });
  it('clamps to the floor and ceiling', () => {
    expect(clampScale(0.01, 0.1, 4)).toBe(0.1);
    expect(clampScale(9, 0.1, 4)).toBe(4);
  });
  it('degenerate bounds resolve to the floor', () => {
    expect(clampScale(1, 4, 0.1)).toBe(4);
  });
});

describe('fitScale', () => {
  const input = { contentWidth: 800, contentHeight: 600, containerWidth: 400, containerHeight: 600 };

  it("'width' fits the width only", () => {
    expect(fitScale(input, 'width')).toBeCloseTo(0.5, 10); // 400/800
  });

  it("'page' fits the tighter of width/height", () => {
    // width fit = 400/800 = 0.5; height fit = 600/600 = 1 → min = 0.5.
    expect(fitScale(input, 'page')).toBeCloseTo(0.5, 10);
    // A short-but-wide container: height now binds.
    expect(
      fitScale({ contentWidth: 800, contentHeight: 600, containerWidth: 1600, containerHeight: 300 }, 'page'),
    ).toBeCloseTo(0.5, 10); // min(2, 0.5)
  });

  it('returns 0 (defer) for an unlaid-out container', () => {
    expect(fitScale({ ...input, containerWidth: 0 }, 'width')).toBe(0);
    expect(fitScale({ ...input, containerHeight: 0 }, 'page')).toBe(0);
    // 'width' mode ignores containerHeight, so a 0 height still fits by width.
    expect(fitScale({ ...input, containerHeight: 0 }, 'width')).toBeCloseTo(0.5, 10);
  });

  it('returns 0 for empty content', () => {
    expect(fitScale({ ...input, contentWidth: 0 }, 'width')).toBe(0);
    expect(fitScale({ ...input, contentHeight: 0 }, 'page')).toBe(0);
  });
});
