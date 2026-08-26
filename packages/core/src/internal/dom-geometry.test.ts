import { describe, expect, it } from 'vitest';
import { intersectElementRects } from './dom-geometry.js';

describe('intersectElementRects', () => {
  it('clips a partially visible card and omits a fully hidden card', () => {
    const viewport = { x: 100, y: 0, width: 280, height: 200 };
    expect(intersectElementRects(
      { x: 100, y: 180, width: 280, height: 60 },
      viewport,
    )).toEqual({ x: 100, y: 180, width: 280, height: 20 });
    expect(intersectElementRects(
      { x: 100, y: 220, width: 280, height: 60 },
      viewport,
    )).toBeUndefined();
  });
});
