import { describe, expect, it } from 'vitest';
import { resolveDocxCommentThreads, type CommentAnchorRange } from './comments.js';
import { clippedBoxVisible, visiblePieces } from './run-clip.js';
import type { DocxClipRegion } from './types.js';

const strips = (from: number, count: number): DocxClipRegion => ({
  x: 0, y: 0, width: 100, height: 20,
  rects: Array.from({ length: count }, (_, index) => ({ x: from + index * 10, y: 0, width: 2, height: 20 })),
});

describe('visibility under clip unions of many rectangles', () => {
  it('shows nothing of a box under two unions that share no area', () => {
    // 9 x 8 rectangle pairs, none of which meet.
    expect(clippedBoxVisible({ x: 0, y: 0, width: 100, height: 20 }, [strips(0, 9), strips(4, 8)])).toBe(false);
  });

  it('shows a box under two such unions that meet in one rectangle', () => {
    const meeting = strips(4, 8);
    const clips = [strips(0, 9), { ...meeting, rects: [...meeting.rects!, { x: 41, y: 0, width: 2, height: 20 }] }];
    expect(clippedBoxVisible({ x: 0, y: 0, width: 100, height: 20 }, clips)).toBe(true);
  });
});

describe('the visible pieces under clip unions of many rectangles', () => {
  const meeting = strips(4, 8);
  const clips = [strips(0, 9), { ...meeting, rects: [...meeting.rects!, { x: 41, y: 0, width: 2, height: 20 }] }];
  const rounded = (value: number) => Math.round(value * 1e6) / 1e6;
  const extent = (piece: readonly Readonly<{ x: number; y: number }>[]) => {
    const xs = piece.map((point) => point.x);
    return [Math.min(...xs), Math.max(...xs)].map(rounded);
  };

  it('are the parts that every union shows, not the rectangles of one of them', () => {
    // The strip at 40..42 of the first union and the one at 41..43 of the
    // second share x 41..42; no other pair of strips meets.
    const pieces = visiblePieces({ x: 0, y: 0, width: 100, height: 20 }, clips);
    expect(pieces.map(extent)).toEqual([[41, 42]]);
  });

  it('give a comment on a run under them one highlight, where the run is seen', () => {
    const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const anchor: CommentAnchorRange = {
      commentId: '1', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    };
    const [thread] = resolveDocxCommentThreads([{ id: '1', text: 'c' }], [anchor], [{
      source, sourceRunIndex: 0, text: 'commented', x: 0, y: 0, w: 100, h: 20,
      fontSize: 16, font: '16px sans-serif', clips,
    }]);
    expect(thread!.anchors[0]!.rects.map((rect) => [rect.x, rect.x + rect.width].map(rounded))).toEqual([[41, 42]]);
  });

  it('are bounded: a union whose rectangles leave too many pieces is cut as its extent', () => {
    // 300 strips cross the box, more pieces than are tracked; every piece
    // returned still has a visible part, and one clip hides the rest.
    const many: DocxClipRegion = {
      x: 0, y: 0, width: 100, height: 20,
      rects: Array.from({ length: 300 }, (_, index) => ({ x: index / 4, y: 0, width: 0.1, height: 20 })),
    };
    const window: DocxClipRegion = { x: 10.02, y: 0, width: 0.03, height: 20 };
    const box = { x: 0, y: 0, width: 100, height: 20 };
    expect(visiblePieces(box, [many]).length).toBeLessThanOrEqual(256);
    expect(visiblePieces(box, [many, window]).map(extent)).toEqual([[10.02, 10.05]]);
    // 10.12..10.15 lies between two strips: inside the extent, not visible.
    expect(visiblePieces(box, [many, { ...window, x: 10.12 }])).toEqual([]);
  });
});
