import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CommentAnchorRange } from './comments.js';
import { buildDocxCommentMargin, resolvePageCommentAnchors } from './comment-margin.js';
import { installDom, makeEl } from './scroll-viewer-test-dom.js';
import type { DocxStorySource, DocxTextRunInfo } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

function source(path: number[]): DocxStorySource {
  return { story: 'body', storyInstance: 'body', path };
}

function anchor(commentId: string, path: number[]): CommentAnchorRange {
  const authored = source(path);
  return {
    commentId,
    source: authored,
    startRunIndex: 0,
    endRunIndex: 1,
    reference: { source: authored, runIndex: 1, affinity: 'preceding' },
  };
}

function run(path: number[]): DocxTextRunInfo {
  return {
    source: source(path), sourceRunIndex: 0, text: 'visible',
    x: 10, y: 20, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
  };
}

describe('resolvePageCommentAnchors', () => {
  it('resolves an authored source and a final-state fallback on the mounted page', () => {
    const direct = anchor('direct', [4]);
    const deleted = {
      ...anchor('deleted', [1]),
      geometryFallback: { source: source([4]), sourceRunIndex: 0 },
    };

    expect(resolvePageCommentAnchors([direct, deleted], [run([4])]).map(({ anchor }) =>
      anchor.commentId)).toEqual(['direct', 'deleted']);
  });

  it('does not resolve thousands of anchors from unrelated page sources', () => {
    let unrelatedRangeReads = 0;
    const unrelated = Array.from({ length: 5_000 }, (_, index) => {
      const item = anchor(`other-${index}`, [index + 10]);
      return Object.defineProperties(item, {
        startRunIndex: { get: () => { unrelatedRangeReads++; return 0; } },
        endRunIndex: { get: () => { unrelatedRangeReads++; return 1; } },
      });
    });
    const visible = anchor('visible', [2]);

    const resolved = resolvePageCommentAnchors([...unrelated, visible], [run([2])]);

    expect(resolved).toHaveLength(1);
    expect(resolved[0]?.anchor.commentId).toBe('visible');
    expect(unrelatedRangeReads).toBe(0);
  });

  it('tints a later range without duplicating the card from the first anchor page', () => {
    installDom();
    const comments = [{ id: 'multi', author: 'Ada', text: 'Across pages' }];
    const anchors = [anchor('multi', [1]), anchor('multi', [2])];
    const render = (path: number[]) => {
      const tint = makeEl('div');
      const margin = makeEl('div');
      margin.clientHeight = 800;
      buildDocxCommentMargin(
        tint as unknown as HTMLDivElement,
        margin as unknown as HTMLDivElement,
        [run(path)],
        { comments, anchors },
        600,
        800,
        null,
        () => undefined,
        1,
        260,
        false,
      );
      return { tint, margin };
    };

    const firstPage = render([1]);
    expect(firstPage.tint.children).toHaveLength(1);
    expect(firstPage.margin.children).toHaveLength(1);

    const laterPage = render([2]);
    expect(laterPage.tint.children).toHaveLength(1);
    expect(laterPage.margin.children).toHaveLength(0);
  });
});
