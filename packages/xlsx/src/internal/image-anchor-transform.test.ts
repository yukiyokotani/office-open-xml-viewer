import { describe, expect, it } from 'vitest';
import { inverseImageTransformPoint, rotatedImageBounds } from './image-anchor-transform.js';

describe('DrawingML picture transform geometry', () => {
  const rect = { x: 10, y: 20, width: 80, height: 30 };

  it.each([
    [0, 80, 30],
    [90, 30, 80],
    [180, 80, 30],
    [30, 80 * Math.cos(Math.PI / 6) + 15, 40 + 30 * Math.cos(Math.PI / 6)],
  ])('computes the rotated culling bounds at %s degrees', (rotation, width, height) => {
    const bounds = rotatedImageBounds(rect, rotation);
    expect(bounds.width).toBeCloseTo(width, 9);
    expect(bounds.height).toBeCloseTo(height, 9);
    expect(bounds.x + bounds.width / 2).toBe(50);
    expect(bounds.y + bounds.height / 2).toBe(35);
  });

  it('inverse-maps asymmetric corners through rotation and both reflections', () => {
    const authored = { x: 18, y: 24 };
    // Forward transform: horizontal + vertical reflections about (50,35), then
    // clockwise 90 degrees. The inverse must recover the exact authored point.
    const painted = { x: 39, y: 67 };
    const recovered = inverseImageTransformPoint(painted, rect, 90, true, true);
    expect(recovered.x).toBeCloseTo(authored.x, 9);
    expect(recovered.y).toBeCloseTo(authored.y, 9);
  });

  it('retains the identity fast path geometry', () => {
    const point = { x: Number.MAX_SAFE_INTEGER - .25, y: -123456789.125 };
    expect(rotatedImageBounds(rect)).toBe(rect);
    expect(inverseImageTransformPoint(point, {
      x: Number.MAX_SAFE_INTEGER - 1000.5, y: -123456999.75, width: 999.5, height: 500.25,
    })).toBe(point);
  });
});
