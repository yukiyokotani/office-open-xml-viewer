import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CommentAnchorRange } from './comments.js';
import { buildDocxCommentMargin, resolvePageCommentAnchors } from './comment-margin.js';
import { resolveDocxCommentThreads } from './comments.js';
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

  it('merges the tints of a cell line across a commented tab into one tint in the shared clip container', () => {
    installDom();
    const comments = [{ id: 'cell', author: 'Ada', text: 'In a cell' }];
    const clips = [{ x: 0, y: 0, width: 300, height: 100 }];
    const tints = (between: DocxTextRunInfo[]) => {
      // Two runs of one comment either side of a 40px gap, in the same cell clip.
      const runs: DocxTextRunInfo[] = [
        { ...run([1]), x: 10, w: 40, clips },
        ...between,
        { ...run([1]), x: 90, w: 40, clips },
      ];
      const tint = makeEl('div');
      buildDocxCommentMargin(tint as unknown as HTMLDivElement, null, runs,
        { comments, anchors: [{ ...anchor('cell', [1]), endRunIndex: 1 }] }, 600, 800, null, () => undefined, 1, 260, false);
      expect(tint.children).toHaveLength(1);
      return tint.children[0]!.children.length;
    };
    // The gap is a tab of the commented run: one tint over text and tab.
    expect(tints([{ ...run([1]), text: '', inlineContent: 'tab', x: 50, w: 40, clips }])).toBe(1);
    // The gap holds content outside the comment: it stays untinted.
    expect(tints([{ ...run([1]), sourceRunIndex: 7, x: 50, w: 40, clips }])).toBe(2);
    expect(tints([])).toBe(2);
  });

  it('uses the same font box and clips for the tint as the comment API', () => {
    installDom();
    const comments = [{ id: 'font', author: 'Ada', text: 'Font box' }];
    const clips = [{ x: 0, y: 0, width: 300, height: 100 }];
    // A 4px exact line whose font box is 7px high.
    const runs: DocxTextRunInfo[] = [{
      ...run([1]), y: 20, h: 4, clips, highlightBounds: { x: 10, y: 18, width: 80, height: 7 },
    }];
    const anchors = [anchor('font', [1])];
    const [thread] = resolveDocxCommentThreads(comments as never, anchors, runs) as unknown as Array<{
      anchors: Array<{ rects: Array<{ x: number; y: number; width: number; height: number; clips?: unknown }> }>;
    }>;
    const rect = thread!.anchors[0]!.rects[0]!;
    expect([rect.x, rect.y, rect.width, rect.height]).toEqual([10, 18, 80, 7]);
    expect(rect.clips).toEqual(clips);
    const tint = makeEl('div');
    buildDocxCommentMargin(tint as unknown as HTMLDivElement, null, runs,
      { comments, anchors }, 600, 800, null, () => undefined, 1, 260, false);
    const [container] = tint.children;
    const [shade] = container!.children;
    // In percentages of the 300 x 100 clip container: the same 7px font box.
    const percent = (name: string) => Number(new RegExp(`${name}:([^%;]+)%`).exec(shade!.style.cssText)?.[1]);
    expect(percent('top')).toBeCloseTo(18, 9);
    expect(percent('height')).toBeCloseTo(7, 9);
  });

  it('joins tints along a turned line and never across two turned columns', () => {
    installDom();
    const comments = [{ id: 'turned', author: 'Ada', text: 'Turned' }];
    const turned = (x: number, y: number): DocxTextRunInfo => ({
      ...run([1]), x, y, w: 40, h: 14, transform: 'rotate(90deg)',
    });
    const tints = (runs: DocxTextRunInfo[]) => {
      const tint = makeEl('div');
      buildDocxCommentMargin(tint as unknown as HTMLDivElement, null, runs,
        { comments, anchors: [anchor('turned', [1])] }, 600, 800, null, () => undefined, 1, 260, false);
      return tint.children.length;
    };
    // Two columns side by side at the same page y: two tints.
    expect(tints([turned(200, 100), turned(150, 100)])).toBe(2);
    // One column, runs painted one after the other down its line: one tint.
    expect(tints([turned(200, 100), turned(200, 140)])).toBe(1);
    // A gap down the line is not tinted.
    expect(tints([turned(200, 100), turned(200, 150)])).toBe(2);
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
