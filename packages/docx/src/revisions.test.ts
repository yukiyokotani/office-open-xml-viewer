import { describe, expect, it } from 'vitest';
import type { LayoutSourceStore } from './layout/layout-source-store.js';
import type { SourceRef } from './layout/types.js';
import { sourceKey } from './layout/source-key.js';
import {
  collectLayoutSourceRevisionRanges,
  resolveRevisionAnchorRuns,
} from './revisions.js';

const sourceRef: SourceRef = { story: 'body', storyInstance: 'body', path: [0] };

function sourceWithRuns(runs: readonly unknown[]): LayoutSourceStore {
  return {
    blocks: {
      sources: [sourceRef],
      resolve: () => ({ type: 'paragraph', runs }),
    },
  } as unknown as LayoutSourceStore;
}

describe('revision anchor projection', () => {
  it('anchors deleted content to the authored following final-state run', () => {
    const source = sourceWithRuns([
      { type: 'text', text: 'old', revision: { kind: 'deletion', id: '7' } },
      { type: 'text', text: 'new', revision: { kind: 'insertion', id: '8' } },
    ]);
    const ranges = collectLayoutSourceRevisionRanges([
      { kind: 'deletion', id: '7', text: 'old' },
      { kind: 'insertion', id: '8', text: 'new' },
    ], source, new Map([[sourceKey(sourceRef), new Set([1])]]));

    expect(ranges).toEqual([
      expect.objectContaining({
        revisionIndex: 0,
        startRunIndex: 0,
        endRunIndex: 1,
        geometryFallback: expect.objectContaining({ sourceRunIndex: 1 }),
      }),
      expect.objectContaining({
        revisionIndex: 1,
        startRunIndex: 1,
        endRunIndex: 2,
      }),
    ]);
    const rendered = [{
      source: sourceRef,
      sourceRunIndex: 1,
      text: 'new',
      x: 10,
      y: 20,
      w: 30,
      h: 12,
      fontSize: 11,
      font: '11px serif',
    }];
    expect(resolveRevisionAnchorRuns(ranges[0]!, rendered).map((run) => run.text)).toEqual(['new']);
    expect(resolveRevisionAnchorRuns(ranges[1]!, rendered).map((run) => run.text)).toEqual(['new']);
  });

  it('does not guess anchors for missing, invalid, or duplicate decimal ids', () => {
    const source = sourceWithRuns([
      { type: 'text', text: 'a', revision: { kind: 'deletion' } },
      { type: 'text', text: 'b', revision: { kind: 'deletion', id: 'not-an-integer' } },
      { type: 'text', text: 'nbsp', revision: { kind: 'deletion', id: '\u00a01\u00a0' } },
      { type: 'text', text: 'c', revision: { kind: 'deletion', id: '1' } },
    ]);
    const ranges = collectLayoutSourceRevisionRanges([
      { kind: 'deletion', text: 'a' },
      { kind: 'deletion', id: 'not-an-integer', text: 'b' },
      { kind: 'deletion', id: '\u00a01\u00a0', text: 'nbsp' },
      { kind: 'deletion', id: '01', text: 'c' },
      { kind: 'deletion', id: '+1', text: 'd' },
    ], source);

    expect(ranges).toEqual([]);
  });
});
