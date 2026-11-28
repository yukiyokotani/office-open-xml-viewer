import { describe, expect, it } from 'vitest';
import { clippedBoxVisible } from './run-clip.js';
import type { DocxClipRegion } from './types.js';

const strips = (from: number, count: number): DocxClipRegion => ({
  x: 0, y: 0, width: 100, height: 20,
  rects: Array.from({ length: count }, (_, index) => ({ x: from + index * 10, y: 0, width: 2, height: 20 })),
});

describe('visibility under clip unions past the piece cap', () => {
  it('shows nothing of a box under two unions that share no area', () => {
    // 9 x 8 rectangle pairs exceed the cap, where the second union is cut
    // as its bounding box; the unions themselves do not meet.
    expect(clippedBoxVisible({ x: 0, y: 0, width: 100, height: 20 }, [strips(0, 9), strips(4, 8)])).toBe(false);
  });

  it('shows a box under two such unions that meet in one rectangle', () => {
    const meeting = strips(4, 8);
    const clips = [strips(0, 9), { ...meeting, rects: [...meeting.rects!, { x: 41, y: 0, width: 2, height: 20 }] }];
    expect(clippedBoxVisible({ x: 0, y: 0, width: 100, height: 20 }, clips)).toBe(true);
  });
});
