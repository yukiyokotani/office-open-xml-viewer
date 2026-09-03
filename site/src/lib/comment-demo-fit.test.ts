import { describe, expect, it } from 'vitest';
import { commentSurfaceFitScale } from './comment-demo-fit.js';

describe('comment demo initial fit', () => {
  it('shrinks a centered page enough to reveal a right-side comment margin', () => {
    expect(commentSurfaceFitScale({
      currentScale: 1,
      viewportWidth: 1000,
      contentWidth: 1000,
      leadingExtent: 0,
      trailingExtent: 300,
    })).toBeCloseTo(0.625, 10);
  });

  it('uses the same geometry for a left-side review margin', () => {
    expect(commentSurfaceFitScale({
      currentScale: 1.5,
      viewportWidth: 900,
      contentWidth: 760,
      leadingExtent: 240,
      trailingExtent: 0,
    })).toBeCloseTo(1.5 * 900 / 1240, 10);
  });

  it('does not enlarge an already visible comment surface', () => {
    expect(commentSurfaceFitScale({
      currentScale: 0.5,
      viewportWidth: 1200,
      contentWidth: 600,
      leadingExtent: 100,
      trailingExtent: 100,
    })).toBeNull();
  });

  it('rejects incomplete or non-finite layout measurements', () => {
    expect(commentSurfaceFitScale({
      currentScale: 1,
      viewportWidth: 0,
      contentWidth: 600,
      leadingExtent: 100,
      trailingExtent: 0,
    })).toBeNull();
    expect(commentSurfaceFitScale({
      currentScale: 1,
      viewportWidth: 1000,
      contentWidth: Number.POSITIVE_INFINITY,
      leadingExtent: 100,
      trailingExtent: 0,
    })).toBeNull();
  });
});
