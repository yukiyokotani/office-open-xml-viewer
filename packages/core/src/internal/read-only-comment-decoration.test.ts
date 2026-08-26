import { describe, expect, it } from 'vitest';
import {
  projectReadOnlyCommentMarginScroll,
  readOnlyCommentConnectorPath,
} from './read-only-comment-decoration.js';

describe('readOnlyCommentConnectorPath', () => {
  const start = { x: 100, y: 40 };
  const end = { x: 240, y: 120 };

  it('builds solid bezier and orthogonal routes', () => {
    expect(readOnlyCommentConnectorPath(start, end, 'bezier')).toContain(' C ');
    const orthogonal = readOnlyCommentConnectorPath(start, end, 'orthogonal');
    expect(orthogonal).toContain(' H ');
    expect(orthogonal).toContain(' V ');
  });

  it('routes toward a left-side card rather than bending right first', () => {
    const path = readOnlyCommentConnectorPath(
      { x: 200, y: 40 },
      { x: 40, y: 120 },
      'bezier',
    );
    expect(path).toContain('C 120 40, 120 120, 40 120');
  });
});

describe('projectReadOnlyCommentMarginScroll', () => {
  it('translates cached card geometry and can reveal a previously clipped card', () => {
    const anchor = Object.freeze({ x: 20, y: 40, width: 80, height: 12 });
    const geometry = Object.freeze({
      threads: Object.freeze([Object.freeze({
        occurrenceKey: 'thread',
        active: false,
        anchorRects: Object.freeze([anchor]),
        cardRect: Object.freeze({ x: 240, y: 220, width: 120, height: 40 }),
      })]),
      cardClipBounds: Object.freeze({ x: 200, y: 100, width: 200, height: 100 }),
      scrollTop: 0,
    });

    expect(projectReadOnlyCommentMarginScroll(geometry, 0)[0]?.cardRect).toBeUndefined();
    expect(projectReadOnlyCommentMarginScroll(geometry, 50)[0]).toEqual({
      occurrenceKey: 'thread',
      active: false,
      anchorRects: [anchor],
      cardRect: { x: 240, y: 170, width: 120, height: 30 },
    });
  });
});
