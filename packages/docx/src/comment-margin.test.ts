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

describe('a comment whose text the page clips away', () => {
  it('gets no tint and no card, while the same comment on visible text gets both', () => {
    installDom();
    const comments = [{ id: 'clipped', author: 'Ada', text: 'Hidden' }];
    const anchors = [anchor('clipped', [1])];
    const render = (clips: DocxTextRunInfo['clips']) => {
      const tint = makeEl('div');
      const margin = makeEl('div');
      margin.clientHeight = 800;
      const runs = [{ ...run([1]), clips }];
      buildDocxCommentMargin(tint as unknown as HTMLDivElement, margin as unknown as HTMLDivElement, runs,
        { comments, anchors }, 600, 800, null, () => undefined, 1, 260, false);
      return { tint, margin, resolved: resolvePageCommentAnchors(anchors, runs) };
    };
    // The run is painted at y=20..34; the row's clip ends above it.
    const hidden = render([{ x: 0, y: 0, width: 300, height: 12 }]);
    expect(hidden.resolved).toEqual([]);
    expect(hidden.tint.children).toHaveLength(0);
    expect(hidden.margin.children).toHaveLength(0);
    const shown = render([{ x: 0, y: 0, width: 300, height: 100 }]);
    expect(shown.resolved).toHaveLength(1);
    expect(shown.margin.children).toHaveLength(1);
  });
});

describe('the card of a comment whose first paragraph the page clips away', () => {
  const comments = [{ id: 'partly', author: 'Ada', text: 'Two paragraphs' }];
  const anchors = [anchor('partly', [1]), anchor('partly', [2])];
  const render = (runs: DocxTextRunInfo[]) => {
    const tint = makeEl('div');
    const margin = makeEl('div');
    margin.clientHeight = 800;
    buildDocxCommentMargin(tint as unknown as HTMLDivElement, margin as unknown as HTMLDivElement, runs,
      { comments, anchors }, 600, 800, null, () => undefined, 1, 260, true);
    const markers = (element: ReturnType<typeof makeEl>): number => element.children.reduce(
      (sum, child) => sum + (child.dataset.ooxmlCommentMarker !== undefined ? 1 : 0) + markers(child), 0);
    return { cards: margin.children.length, markers: markers(tint) };
  };

  it('stands with the first paragraph a reader can see', () => {
    installDom();
    // The first paragraph is painted outside its cell's clip; the second is not clipped.
    const hidden = { ...run([1]), clips: [{ x: 200, y: 0, width: 100, height: 100 }] };
    const shown = { ...run([2]), y: 60 };
    expect(render([hidden, shown])).toEqual({ cards: 1, markers: 1 });
  });

  it('stays with the first paragraph when that is on another page', () => {
    installDom();
    expect(render([{ ...run([2]), y: 60 }])).toEqual({ cards: 0, markers: 0 });
  });
});
