import { describe, expect, it } from 'vitest';
import {
  collectLayoutSourceCommentRanges,
  collectLayoutSourceCommentRangesIfPresent,
  collectStoryCommentRanges,
  resolveCommentAnchorRuns,
} from './comments.js';
import type { LayoutSourceStore } from './layout/layout-source-store.js';
import type { SourceRef } from './layout/types.js';
import type { DocParagraph, DocxCommentMark, DocxTextRunInfo } from './types.js';

function para(runTexts: readonly string[], marks?: readonly DocxCommentMark[]): DocParagraph {
  return {
    type: 'paragraph',
    runs: runTexts.map((text) => ({ type: 'text', text })),
    ...(marks ? { commentMarks: marks } : {}),
  } as unknown as DocParagraph;
}

function source(
  path: readonly number[],
  story: SourceRef['story'] = 'body',
  storyInstance: string = story,
): SourceRef {
  return { story, storyInstance, path };
}

const mark = (id: string, kind: DocxCommentMark['kind'], runIndex: number): DocxCommentMark =>
  ({ id, kind, runIndex });

const valid = (...ids: string[]): ReadonlySet<string> => new Set(ids);
const reference = (
  path: readonly number[],
  runIndex: number,
  affinity: 'following' | 'preceding',
  story: SourceRef['story'] = 'body',
  storyInstance: string = story,
) => ({ source: source(path, story, storyInstance), runIndex, affinity });

describe('comment data projections', () => {
  it('resolves same- and cross-paragraph ranges within one story', () => {
    expect(collectStoryCommentRanges([
      {
        paragraph: para(['before', 'annotated'], [mark('9', 'rangeStart', 1)]),
        source: source([0]),
      },
      {
        paragraph: para(['continued', 'after'], [
          mark('9', 'rangeEnd', 1),
          mark('9', 'reference', 2),
        ]),
        source: source([1]),
      },
    ], valid('9'))).toEqual([
      {
        commentId: '9', source: source([0]), startRunIndex: 1, endRunIndex: 2,
        reference: reference([1], 2, 'preceding'),
      },
      {
        commentId: '9', source: source([1]), startRunIndex: 0, endRunIndex: 1,
        reference: reference([1], 2, 'preceding'),
      },
    ]);
  });

  it('emits reference-only zero-length anchors', () => {
    expect(collectStoryCommentRanges([{
      paragraph: para(['a', 'b'], [mark('3', 'reference', 1)]),
      source: source([0]),
    }], valid('3'))).toEqual([
      {
        commentId: '3', source: source([0]), startRunIndex: 1, endRunIndex: 1,
        reference: reference([0], 1, 'following'),
      },
    ]);
  });

  it.each([
    ['start', mark('4', 'rangeStart', 1)],
    ['end', mark('4', 'rangeEnd', 1)],
  ])('treats a lone range %s as the spec-defined point anchor', (_kind, boundary) => {
    expect(collectStoryCommentRanges([{
      paragraph: para(['a', 'b'], [boundary, mark('4', 'reference', 2)]),
      source: source([0]),
    }], valid('4'))).toEqual([
      {
        commentId: '4', source: source([0]), startRunIndex: 1, endRunIndex: 1,
        reference: reference([0], 2, 'preceding'),
      },
    ]);
  });

  it('does not pair identical comment IDs across document stories', () => {
    expect([
      ...collectStoryCommentRanges([{
        paragraph: para(['body'], [mark('7', 'rangeStart', 0), mark('7', 'reference', 1)]),
        source: source([0]),
      }], valid('7')),
      ...collectStoryCommentRanges([{
        paragraph: para(['header'], [mark('7', 'rangeEnd', 1), mark('7', 'reference', 1)]),
        source: source([0], 'header', 'default'),
      }], valid('7')),
    ]).toEqual([
      {
        commentId: '7', source: source([0]), startRunIndex: 0, endRunIndex: 0,
        reference: reference([0], 1, 'preceding'),
      },
      {
        commentId: '7',
        source: source([0], 'header', 'default'),
        startRunIndex: 1,
        endRunIndex: 1,
        reference: reference([0], 1, 'preceding', 'header', 'default'),
      },
    ]);
  });

  it('preserves body, header, footer, note, and text-box source identities', () => {
    const sources: SourceRef[] = [
      source([0]),
      source([0], 'header', 'default'),
      source([0], 'footer', 'even'),
      source([0], 'footnote', '4'),
      source([0], 'endnote', '5'),
      source([0], 'textbox', 'body:body:0.1'),
    ];
    const paragraphs = new Map(sources.map((item, index) => [
      `${item.story}:${item.storyInstance}`,
      para([item.story], [mark(String(index), 'reference', 0)]),
    ]));
    const layoutSource = {
      blocks: {
        sources,
        resolve(item: SourceRef) {
          return paragraphs.get(`${item.story}:${item.storyInstance}`) as DocParagraph;
        },
      },
    } as unknown as LayoutSourceStore;

    const comments = sources.map((_item, index) => ({ id: String(index) }));
    expect(collectLayoutSourceCommentRanges(comments, layoutSource)).toEqual(
      sources.map((item, index) => ({
        commentId: String(index),
        source: item,
        startRunIndex: 0,
        endRunIndex: 0,
        reference: {
          source: item,
          runIndex: 0,
          affinity: 'following',
        },
      })),
    );
  });

  it.each([
    ['missing comment', [mark('8', 'reference', 0)], valid('9')],
    ['non-decimal id', [mark('x', 'reference', 0)], valid('x')],
    ['missing reference', [mark('8', 'rangeStart', 0), mark('8', 'rangeEnd', 1)], valid('8')],
    ['duplicate reference', [mark('8', 'reference', 0), mark('8', 'reference', 1)], valid('8')],
    ['duplicate start', [
      mark('8', 'rangeStart', 0), mark('8', 'rangeStart', 1), mark('8', 'reference', 1),
    ], valid('8')],
    ['reversed pair', [
      mark('8', 'rangeStart', 1), mark('8', 'rangeEnd', 0), mark('8', 'reference', 1),
    ], valid('8')],
  ])('fails closed for malformed marks: %s', (_case, marks, ids) => {
    expect(collectStoryCommentRanges([{
      paragraph: para(['a'], marks as DocxCommentMark[]),
      source: source([0]),
    }], ids as ReadonlySet<string>)).toEqual([]);
  });

  it.each([
    ['+1', '1', '1'],
    [' 01 ', '+1', '+1'],
    ['-0', '0', '0'],
    ['00000000000000000000000042', '+42', '+42'],
  ])('matches ST_DecimalNumber value-space IDs %s and %s', (markId, commentId, outputId) => {
    expect(collectStoryCommentRanges([{
      paragraph: para(['value'], [mark(markId, 'reference', 0)]),
      source: source([0]),
    }], valid(commentId))).toEqual([{
      commentId: outputId,
      source: source([0]),
      startRunIndex: 0,
      endRunIndex: 0,
      reference: reference([0], 0, 'following'),
    }]);
  });

  it('fails closed when two comment IDs have the same integer value', () => {
    expect(collectStoryCommentRanges([{
      paragraph: para(['value'], [mark('1', 'reference', 0)]),
      source: source([0]),
    }], valid('1', '01'))).toEqual([]);
  });

  it('does not treat non-XML Unicode whitespace as ST_DecimalNumber collapse space', () => {
    expect(collectStoryCommentRanges([{
      paragraph: para(['value'], [mark('\u00a01\u00a0', 'reference', 0)]),
      source: source([0]),
    }], valid('1'))).toEqual([]);
  });

  it('publishes an adjacent visible-run fallback for a deleted-only paragraph', () => {
    const deleted = {
      type: 'paragraph',
      runs: [{ type: 'text', text: 'gone', revision: { kind: 'deletion' } }],
      commentMarks: [mark('4', 'reference', 1)],
    } as unknown as DocParagraph;
    const rendered = [{
      source: source([1]), sourceRunIndex: 0, text: 'visible',
    }] as unknown as DocxTextRunInfo[];
    expect(collectStoryCommentRanges([
      { paragraph: deleted, source: source([0]) },
      { paragraph: para(['visible']), source: source([1]) },
    ], valid('4'), rendered)).toEqual([{
      commentId: '4',
      source: source([0]),
      startRunIndex: 1,
      endRunIndex: 1,
      reference: reference([0], 1, 'preceding'),
      geometryFallback: { source: source([1]), sourceRunIndex: 0 },
    }]);
  });

  it('derives cross-paragraph fallback from actual layout geometry, not field kind', () => {
    const deleted = {
      type: 'paragraph',
      runs: [{ type: 'text', text: 'gone', revision: { kind: 'deletion' } }],
      commentMarks: [mark('4', 'reference', 1)],
    } as unknown as DocParagraph;
    const emptyField = {
      type: 'paragraph',
      runs: [{ type: 'field', instr: 'UNKNOWN', fallbackText: '' }],
    } as unknown as DocParagraph;
    const rendered = [{
      source: source([2]), sourceRunIndex: 0, text: 'visible',
    }] as unknown as DocxTextRunInfo[];
    expect(collectStoryCommentRanges([
      { paragraph: deleted, source: source([0]) },
      { paragraph: emptyField, source: source([1]) },
      { paragraph: para(['visible']), source: source([2]) },
    ], valid('4'), rendered)).toEqual([expect.objectContaining({
      geometryFallback: { source: source([2]), sourceRunIndex: 0 },
    })]);
  });

  it('resolves covered, gapped, split, shuffled, and cross-paragraph geometry', () => {
    const run = (path: readonly number[], sourceRunIndex: number, text: string) => ({
      source: source(path), sourceRunIndex, text,
    }) as unknown as DocxTextRunInfo;
    const shuffled = [
      run([0], 5, 'five'),
      run([0], 1, 'one'),
      run([0], 4, 'four-a'),
      run([0], 4, 'four-b'),
      run([1], 0, 'fallback'),
    ];
    const point = {
      commentId: '1', source: source([0]), startRunIndex: 3, endRunIndex: 3,
      reference: reference([0], 3, 'following'),
    };
    expect(resolveCommentAnchorRuns(point, shuffled).map(({ text }) => text))
      .toEqual(['four-a', 'four-b']);
    expect(resolveCommentAnchorRuns({
      ...point,
      reference: reference([0], 3, 'preceding'),
    }, shuffled).map(({ text }) => text)).toEqual(['four-a', 'four-b']);
    expect(resolveCommentAnchorRuns({
      ...point,
      startRunIndex: 2,
      endRunIndex: 3,
      reference: reference([0], 3, 'preceding'),
    }, shuffled).map(({ text }) => text)).toEqual(['one']);
    expect(resolveCommentAnchorRuns({
      ...point,
      source: source([2]),
      reference: reference([2], 0, 'following'),
      geometryFallback: { source: source([1]), sourceRunIndex: 0 },
    }, shuffled).map(({ text }) => text)).toEqual(['fallback']);
    expect(resolveCommentAnchorRuns({
      ...point,
      startRunIndex: 4,
      endRunIndex: 6,
    }, shuffled).map(({ text }) => text)).toEqual(['five', 'four-a', 'four-b']);
  });

  it('resolves a lone range point at its own boundary rather than a distant reference', () => {
    const runs = [
      { source: source([0]), sourceRunIndex: 1, text: 'point' },
      { source: source([0]), sourceRunIndex: 5, text: 'reference' },
    ] as unknown as DocxTextRunInfo[];
    expect(resolveCommentAnchorRuns({
      commentId: '1',
      source: source([0]),
      startRunIndex: 1,
      endRunIndex: 1,
      reference: reference([0], 5, 'preceding'),
    }, runs).map(({ text }) => text)).toEqual(['point']);
  });

  it('does not walk retained stories when no comments exist', () => {
    const layoutSource = new Proxy({} as LayoutSourceStore, {
      get() {
        throw new Error('layout source must not be walked');
      },
    });
    expect(collectLayoutSourceCommentRangesIfPresent([], layoutSource)).toEqual([]);
  });
});
