import { describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import type { CommentAnchorRange } from './comments.js';
import type { DocComment, DocxTextRunInfo } from './types.js';

describe('DocxDocument.getCommentThreads', () => {
  it('returns only threads rendered on the requested page', async () => {
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    const comments: DocComment[] = [
      { id: '1', author: 'Ada', text: 'Visible root' },
      { id: '2', parentId: '1', author: 'Bob', text: 'Reply' },
      { id: '3', author: 'Chen', text: 'Other page' },
    ];
    const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const anchors: CommentAnchorRange[] = [
      {
        commentId: '1',
        source,
        startRunIndex: 0,
        endRunIndex: 1,
        reference: { source, runIndex: 1, affinity: 'preceding' },
      },
      {
        commentId: '3',
        source: { story: 'body', storyInstance: 'body', path: [1] },
        startRunIndex: 0,
        endRunIndex: 1,
        reference: {
          source: { story: 'body', storyInstance: 'body', path: [1] },
          runIndex: 1,
          affinity: 'preceding',
        },
      },
    ];
    const runs: DocxTextRunInfo[] = [{
      source,
      sourceRunIndex: 0,
      text: 'Visible',
      x: 10,
      y: 20,
      w: 30,
      h: 12,
      fontSize: 12,
      font: '12px sans-serif',
    }];

    Object.defineProperty(document, 'comments', { value: comments });
    document.commentAnchorRanges = vi.fn(() => anchors);
    document.collectPageRuns = vi.fn(async () => runs);

    const threads = await document.getCommentThreads(2, {
      width: 720,
      includeResolved: false,
    });

    expect(document.collectPageRuns).toHaveBeenCalledWith(2, {
      width: 720,
      currentDate: undefined,
    });
    expect(threads).toHaveLength(1);
    expect(threads[0]?.root.id).toBe('1');
    expect(threads[0]?.replies.map(({ id }) => id)).toEqual(['2']);
    expect(threads[0]?.anchors[0]?.rects).toEqual([
      { x: 10, y: 20, width: 30, height: 12 },
    ]);
  });

  it('does not scan page runs once per off-page anchor', async () => {
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    const visibleSource = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const offPageAnchors: CommentAnchorRange[] = Array.from({ length: 5_000 }, (_, index) => {
      const source = { story: 'body' as const, storyInstance: 'body', path: [index + 1] };
      return {
        commentId: `off-page-${index}`,
        source,
        startRunIndex: 0,
        endRunIndex: 1,
        reference: { source, runIndex: 1, affinity: 'preceding' },
      };
    });
    const visibleAnchor: CommentAnchorRange = {
      commentId: 'visible',
      source: visibleSource,
      startRunIndex: 0,
      endRunIndex: 1,
      reference: { source: visibleSource, runIndex: 1, affinity: 'preceding' },
    };
    let sourceReads = 0;
    const run = {
      get source() {
        sourceReads += 1;
        return visibleSource;
      },
      sourceRunIndex: 0,
      text: 'Visible',
      x: 10,
      y: 20,
      w: 30,
      h: 12,
      fontSize: 12,
      font: '12px sans-serif',
    } as DocxTextRunInfo;

    Object.defineProperty(document, 'comments', {
      value: [
        ...offPageAnchors.map((anchor) => ({ id: anchor.commentId, text: 'Off page' })),
        { id: 'visible', text: 'Visible' },
      ] satisfies DocComment[],
    });
    document.commentAnchorRanges = vi.fn(() => [...offPageAnchors, visibleAnchor]);
    document.collectPageRuns = vi.fn(async () => [run]);

    const threads = await document.getCommentThreads(0);

    expect(threads.map(({ root }) => root.id)).toEqual(['visible']);
    expect(sourceReads).toBeLessThan(20);
  });
});
