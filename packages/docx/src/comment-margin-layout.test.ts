import { describe, it, expect } from 'vitest';
import {
  buildCommentThreads,
  collectDocumentCommentRanges,
  computeCommentBalloonLayout,
  type CommentBalloonRequest,
} from './comment-margin-layout.js';
import type { BodyElement, DocComment, DocParagraph, DocxCommentMark } from './types';

// Pure comment-margin model (§17.13.4 anchors + commentsExtended threading +
// the balloon stacking conventions): threads, document-order anchor ranges,
// and the gutter placement algorithm, each in isolation.

function comment(id: string, extra: Partial<DocComment> = {}): DocComment {
  return { id, author: `author-${id}`, text: `text-${id}`, ...extra };
}

function para(
  runTexts: readonly string[],
  commentMarks?: readonly DocxCommentMark[],
  runTypes?: readonly string[],
): BodyElement {
  const p = {
    type: 'paragraph',
    alignment: 'left',
    runs: runTexts.map((text, index) => ({
      type: runTypes?.[index] ?? 'text',
      text,
    })),
    ...(commentMarks ? { commentMarks } : {}),
  } as unknown as DocParagraph & { type: 'paragraph' };
  return p as BodyElement;
}

function table(cells: readonly BodyElement[][]): BodyElement {
  return {
    type: 'table',
    rows: [{ cells: cells.map((content) => ({ content })) }],
  } as unknown as BodyElement;
}

describe('buildCommentThreads', () => {
  it('groups replies under their root in comments-part order and drops resolved threads', () => {
    const threads = buildCommentThreads([
      comment('1'),
      comment('2', { parentId: '1' }),
      comment('3', { parentId: '2' }),
      comment('4', { resolved: true }),
      comment('5', { parentId: '4' }),
      comment('6'),
    ]);
    expect(threads.map((thread) => thread.root.id)).toEqual(['1', '6']);
    expect(threads[0]!.replies.map((reply) => reply.id)).toEqual(['2', '3']);
    expect(threads[1]!.replies).toEqual([]);
  });

  it('drops orphan replies and survives a malformed parent cycle', () => {
    const threads = buildCommentThreads([
      comment('1'),
      comment('2', { parentId: 'missing' }),
      comment('3', { parentId: '4' }),
      comment('4', { parentId: '3' }),
    ]);
    expect(threads.map((thread) => thread.root.id)).toEqual(['1']);
    expect(threads[0]!.replies).toEqual([]);
  });
});

describe('collectDocumentCommentRanges', () => {
  const mark = (id: string, kind: DocxCommentMark['kind'], runIndex: number): DocxCommentMark =>
    ({ id, kind, runIndex });

  it('resolves a same-paragraph range to a run interval', () => {
    const ranges = collectDocumentCommentRanges([
      para(['before', 'annotated', 'after'], [
        mark('9', 'rangeStart', 1),
        mark('9', 'rangeEnd', 2),
        mark('9', 'reference', 2),
      ]),
    ]);
    expect(ranges).toEqual([
      { commentId: '9', paragraphPath: [0], startRunIndex: 1, endRunIndex: 2 },
    ]);
  });

  it('carries a cross-paragraph range through mark-less middle paragraphs', () => {
    const ranges = collectDocumentCommentRanges([
      para(['a', 'b'], [mark('1', 'rangeStart', 1)]),
      para(['middle']),
      para(['c', 'd'], [mark('1', 'rangeEnd', 1)]),
    ]);
    expect(ranges).toEqual([
      { commentId: '1', paragraphPath: [0], startRunIndex: 1, endRunIndex: 2 },
      { commentId: '1', paragraphPath: [1], startRunIndex: 0, endRunIndex: 1 },
      { commentId: '1', paragraphPath: [2], startRunIndex: 0, endRunIndex: 1 },
    ]);
  });

  it('addresses table-cell paragraphs with the structural path scheme', () => {
    const ranges = collectDocumentCommentRanges([
      para(['first']),
      table([
        [para(['plain'])],
        [para(['x', 'y'], [mark('7', 'rangeStart', 0), mark('7', 'rangeEnd', 1)])],
      ]),
    ]);
    expect(ranges).toEqual([
      { commentId: '7', paragraphPath: [1, 0, 1, 0], startRunIndex: 0, endRunIndex: 1 },
    ]);
  });

  it('emits a zero-length boundary for a reference-only comment', () => {
    const ranges = collectDocumentCommentRanges([
      para(['a', 'b'], [mark('3', 'reference', 1)]),
    ]);
    expect(ranges).toEqual([
      { commentId: '3', paragraphPath: [0], startRunIndex: 1, endRunIndex: 1 },
    ]);
  });

  it('skips parser-only sidecar runs when normalizing boundaries', () => {
    const ranges = collectDocumentCommentRanges([
      para(
        ['sidecar', 'a', 'b'],
        [mark('5', 'rangeStart', 1), mark('5', 'rangeEnd', 3)],
        ['unavailableDrawing', 'text', 'text'],
      ),
    ]);
    expect(ranges).toEqual([
      { commentId: '5', paragraphPath: [0], startRunIndex: 0, endRunIndex: 2 },
    ]);
  });

  it('keeps an unmatched rangeStart open to the end of the body', () => {
    const ranges = collectDocumentCommentRanges([
      para(['a', 'b'], [mark('2', 'rangeStart', 1)]),
      para(['tail']),
    ]);
    expect(ranges).toEqual([
      { commentId: '2', paragraphPath: [0], startRunIndex: 1, endRunIndex: 2 },
      { commentId: '2', paragraphPath: [1], startRunIndex: 0, endRunIndex: 1 },
    ]);
  });
});

describe('computeCommentBalloonLayout', () => {
  const GEOM = { pageHeightPx: 1000, lineHeightPx: 10, headerHeightPx: 20, gapPx: 5 };
  const balloon = (
    commentId: string,
    anchorYPx: number,
    contentLines: number,
    selected = false,
  ): CommentBalloonRequest => ({ commentId, anchorYPx, contentLines, selected });

  it('places non-competing balloons at their anchors with full capped height', () => {
    const placed = computeCommentBalloonLayout({
      ...GEOM,
      balloons: [balloon('a', 100, 3), balloon('b', 400, 25)],
    });
    expect(placed).toEqual([
      expect.objectContaining({ commentId: 'a', yPx: 100, heightPx: 20 + 3 * 10, visibleLines: 3 }),
      // 25 content lines cap at 10.
      expect.objectContaining({ commentId: 'b', yPx: 400, heightPx: 20 + 10 * 10, visibleLines: 10 }),
    ]);
  });

  it('keeps anchor order and never overlaps (monotonic push-down)', () => {
    const placed = computeCommentBalloonLayout({
      ...GEOM,
      // Same anchor line: the second must start below the first.
      balloons: [balloon('b', 100, 2), balloon('a', 100, 2)],
    });
    expect(placed.map((p) => p.commentId)).toEqual(['b', 'a']);
    expect(placed[1]!.yPx).toBeGreaterThanOrEqual(placed[0]!.yPx + placed[0]!.heightPx + GEOM.gapPx);
  });

  it('truncates a long predecessor so the next balloon reaches its anchor', () => {
    const placed = computeCommentBalloonLayout({
      ...GEOM,
      // Full first balloon (10 lines = 120px tall at y=100) would push the
      // second (anchor 150) down to 225; truncation reclaims the space.
      balloons: [balloon('long', 100, 10), balloon('next', 150, 2)],
    });
    const [long, next] = placed;
    expect(next!.yPx).toBe(150);
    expect(long!.yPx).toBe(100);
    expect(long!.yPx + long!.heightPx + GEOM.gapPx).toBeLessThanOrEqual(next!.yPx);
    expect(long!.visibleLines).toBeGreaterThanOrEqual(1);
    expect(long!.visibleLines).toBeLessThan(10);
  });

  it('never truncates below the one-line floor', () => {
    const placed = computeCommentBalloonLayout({
      ...GEOM,
      // Anchors 5px apart: no amount of truncation lets the second reach its
      // anchor; the first stops at header + one line.
      balloons: [balloon('first', 100, 10), balloon('second', 105, 2)],
    });
    expect(placed[0]!.heightPx).toBe(GEOM.headerHeightPx + GEOM.lineHeightPx);
    expect(placed[0]!.visibleLines).toBe(1);
    expect(placed[1]!.yPx).toBe(placed[0]!.yPx + placed[0]!.heightPx + GEOM.gapPx);
  });

  it('collapses trailing unselected balloons to header stubs under page pressure', () => {
    const placed = computeCommentBalloonLayout({
      ...GEOM,
      pageHeightPx: 100,
      balloons: [balloon('a', 0, 10), balloon('b', 0, 10), balloon('c', 0, 10)],
    });
    expect(placed.some((p) => p.collapsed)).toBe(true);
    const collapsed = placed.filter((p) => p.collapsed);
    for (const stub of collapsed) {
      expect(stub.heightPx).toBe(GEOM.headerHeightPx);
      expect(stub.visibleLines).toBe(0);
    }
    // Collapse eats from the trailing end first.
    expect(placed[placed.length - 1]!.collapsed).toBe(true);
  });

  it('the selected balloon always keeps its full capped height', () => {
    const placed = computeCommentBalloonLayout({
      ...GEOM,
      pageHeightPx: 150,
      balloons: [
        balloon('a', 0, 10),
        balloon('sel', 0, 25, true),
        balloon('c', 0, 10),
      ],
    });
    const selected = placed.find((p) => p.commentId === 'sel')!;
    expect(selected.selected).toBe(true);
    expect(selected.collapsed).toBe(false);
    // Selection raises the cap to the page's line capacity:
    // floor((150 − 20) / 10) = 13 lines (> the shared maxLines 10).
    expect(selected.heightPx).toBe(GEOM.headerHeightPx + 13 * GEOM.lineHeightPx);
    expect(selected.visibleLines).toBe(13);
    // Layout stays deterministic and ordered.
    expect(placed.map((p) => p.commentId)).toEqual(['a', 'sel', 'c']);
    for (let index = 1; index < placed.length; index += 1) {
      expect(placed[index]!.yPx).toBeGreaterThanOrEqual(
        placed[index - 1]!.yPx + placed[index - 1]!.heightPx + GEOM.gapPx,
      );
    }
  });

  it('selection expands past maxLines but never past the page line capacity', () => {
    // 40 content lines: unselected would cap at 10; selected shows all 40
    // (page capacity floor((1000 − 20)/10) = 98 is not the binding cap here).
    const grown = computeCommentBalloonLayout({
      ...GEOM,
      balloons: [balloon('sel', 100, 40, true)],
    });
    expect(grown[0]!.heightPx).toBe(GEOM.headerHeightPx + 40 * GEOM.lineHeightPx);
    expect(grown[0]!.visibleLines).toBe(40);
    // 200 content lines clamp at the page capacity (98 lines fills the page
    // exactly: 20 + 98×10 = 1000); the DOM layer scrolls the rest.
    const clamped = computeCommentBalloonLayout({
      ...GEOM,
      balloons: [balloon('sel', 0, 200, true)],
    });
    expect(clamped[0]!.heightPx).toBe(GEOM.pageHeightPx);
    expect(clamped[0]!.visibleLines).toBe(98);
    // An UNSELECTED balloon with the same content keeps the shared cap.
    const unselected = computeCommentBalloonLayout({
      ...GEOM,
      balloons: [balloon('plain', 100, 40)],
    });
    expect(unselected[0]!.visibleLines).toBe(10);
  });

  it('is deterministic for identical input', () => {
    const input = {
      ...GEOM,
      balloons: [balloon('a', 10, 4), balloon('b', 10, 4), balloon('c', 300, 1)],
    };
    expect(computeCommentBalloonLayout(input)).toEqual(computeCommentBalloonLayout(input));
  });
});
