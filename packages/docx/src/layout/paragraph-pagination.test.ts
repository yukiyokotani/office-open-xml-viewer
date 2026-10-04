import { describe, expect, it } from 'vitest';
import { selectParagraphFragment } from './paragraph-pagination.js';
import type { LineLayout, ParagraphLayout } from './types.js';

const splittable = (lineEndBoundaries: readonly { segIndex: number; charOffset: number }[]) => ({
  kind: 'splittable' as const,
  lineEndBoundaries,
});

const paragraph = (): ParagraphLayout => ({
  kind: 'paragraph', id: 'p', source: { story: 'body', storyInstance: 'body', path: [0] },
  flowDomainId: 'body', ordinaryFlow: true,
  flowBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 30 },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 30 }, advancePt: 30,
  spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
  lines: Array.from({ length: 3 }, (_, index) => ({
    range: { start: index, end: index + 1 },
    bounds: { xPt: 0, yPt: index * 10, widthPt: 100, heightPt: 10 },
    baselinePt: index * 10 + 8, advancePt: 10, placements: [],
  })) as LineLayout[],
  borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
});

const twoLineParagraph = (
  authoredSpaceAfterPt: number,
  retainedTrailingExtentPt: number = authoredSpaceAfterPt,
): ParagraphLayout => ({
  ...paragraph(),
  flowBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 20 + retainedTrailingExtentPt },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 20 },
  advancePt: 20 + retainedTrailingExtentPt,
  spacing: { beforePt: 0, afterPt: retainedTrailingExtentPt },
  lines: paragraph().lines.slice(0, 2),
});

const twoLineBoundaries = [
  { segIndex: 0, charOffset: 1 },
  { segIndex: 0, charOffset: 2 },
];

const verticalEdgeParagraph = (
  writingMode: 'horizontal-tb' | 'vertical-rl',
  baselinePt: number,
  visibleResourceEndPt?: number,
  visibleInkDescentPt?: number,
  blockAxisInkEndPt?: number,
  transformedGlyph = false,
): ParagraphLayout => ({
  ...paragraph(),
  flowBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 25 },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 25 },
  advancePt: 25,
  lines: [{
    range: { start: 0, end: 1 },
    bounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 25 },
    baselinePt,
    advancePt: 25,
    placements: [
      {
        kind: 'text', text: 'A', range: { start: 0, end: 1 },
        origin: { xPt: 0, yPt: baselinePt },
        bounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 25 },
        advancePt: 10, clusters: [], paintOps: visibleInkDescentPt === undefined ? [] : [{
          text: 'A', range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 },
          letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto',
          writingMode,
          ...(transformedGlyph ? { glyphOrientation: 'upright' as const } : {}),
          inkBounds: {
            xMinPt: 0, xMaxPt: 7, ascentPt: 7, descentPt: visibleInkDescentPt,
          },
          ...(blockAxisInkEndPt === undefined ? {} : {
            blockAxisInkBounds: { startPt: -7, endPt: blockAxisInkEndPt },
          }),
        }],
        color: { kind: 'default' },
        fontRoute: { familyList: 'serif', scope: 'native', fingerprint: 'serif' },
        fontSizePt: 10, fontWeight: 400, fontStyle: 'normal',
        direction: 'ltr', writingMode, decorations: [],
      },
      ...(visibleResourceEndPt === undefined ? [] : [{
        kind: 'resource' as const,
        range: { start: 1, end: 2 },
        resourceKey: 'visible-resource', resourceKind: 'image' as const,
        bounds: { xPt: 10, yPt: visibleResourceEndPt - 1, widthPt: 1, heightPt: 1 },
        advancePt: 1,
      }]),
    ],
  }] as LineLayout[],
});

describe('paragraph page-local reserve selection', () => {
  it('does not require a source boundary for a single retained marker-only line', () => {
    const retained = {
      ...paragraph(),
      flowBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 10 },
      inkBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 10 },
      advancePt: 10,
      lines: paragraph().lines.slice(0, 1),
    };

    const selected = selectParagraphFragment(
      retained,
      { boundary: null },
      { kind: 'indivisible' },
      20,
      40,
      true,
      { keepLines: false, widowControl: false },
    );

    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.nextCursor).toBeNull();
  });

  it('fails closed when a splittable retained paragraph lacks a continuation boundary', () => {
    expect(() => selectParagraphFragment(
      paragraph(),
      { boundary: null },
      splittable([{ segIndex: 0, charOffset: 1 }]),
      20,
      40,
      true,
      { keepLines: false, widowControl: false },
    )).toThrow(/align with retained lines/i);
  });

  it('admits a vertical-rl final visible baseline when only the retained line box crosses the block end', () => {
    const selected = selectParagraphFragment(
      verticalEdgeParagraph('vertical-rl', 17.5),
      { boundary: null }, splittable([{ segIndex: 0, charOffset: 1 }]),
      20, 40, true,
      { keepLines: false, widowControl: false, writingMode: 'vertical-rl' },
    );

    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.fragment?.advancePt).toBe(25);
    expect(selected).toMatchObject({
      nextCursor: null,
      requiresFreshFlowRegion: false,
      admittedBlockExtentPt: 20,
    });
  });

  it.each([
    ['horizontal final line', 'horizontal-tb', 17.5, undefined],
    ['vertical final baseline beyond the edge', 'vertical-rl', 21, undefined],
    ['vertical final retained resource ink beyond the edge', 'vertical-rl', 17.5, 21],
  ] as const)('relocates a %s', (_label, writingMode, baselinePt, resourceEndPt) => {
    const selected = selectParagraphFragment(
      verticalEdgeParagraph(writingMode, baselinePt, resourceEndPt),
      { boundary: null }, splittable([{ segIndex: 0, charOffset: 1 }]),
      20, 40, true,
      { keepLines: false, widowControl: false, writingMode },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('relocates a vertical final line whose selected-face ink crosses the edge', () => {
    const selected = selectParagraphFragment(
      verticalEdgeParagraph('vertical-rl', 17.5, undefined, 0, 3, true),
      { boundary: null }, splittable([{ segIndex: 0, charOffset: 1 }]),
      20, 40, true,
      { keepLines: false, widowControl: false, writingMode: 'vertical-rl' },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('fails closed when transformed final-line ink metrics are unavailable', () => {
    const selected = selectParagraphFragment(
      verticalEdgeParagraph('vertical-rl', 17.5, undefined, 0, undefined, true),
      { boundary: null }, splittable([{ segIndex: 0, charOffset: 1 }]),
      20, 40, true,
      { keepLines: false, widowControl: false, writingMode: 'vertical-rl' },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('relocates when a retained decoration stroke crosses the block edge', () => {
    const retained = verticalEdgeParagraph('vertical-rl', 17.5);
    const text = retained.lines[0]!.placements[0]!;
    if (text.kind !== 'text') throw new Error('Expected a text fixture');
    const withStroke: ParagraphLayout = {
      ...retained,
      lines: [{
        ...retained.lines[0]!,
        placements: [{
          ...text,
          decorations: [{
            kind: 'underline', from: { xPt: 0, yPt: 19 }, to: { xPt: 10, yPt: 19 },
            color: '#000', widthPt: 4, style: 'solid',
          }],
        }],
      }],
    };

    const selected = selectParagraphFragment(
      withStroke,
      { boundary: null }, splittable([{ segIndex: 0, charOffset: 1 }]),
      20, 40, true,
      { keepLines: false, widowControl: false, writingMode: 'vertical-rl' },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('admits final visible content when only authored spaceAfter crosses the region edge', () => {
    const selected = selectParagraphFragment(
      twoLineParagraph(10), { boundary: null }, splittable(twoLineBoundaries),
      20, 40, true,
      { keepLines: false, widowControl: true, authoredSpaceAfterPt: 10 },
    );

    expect(selected.fragment?.lines).toHaveLength(2);
    expect(selected.fragment?.advancePt).toBe(30);
    expect(selected).toMatchObject({
      nextCursor: null,
      requiresFreshFlowRegion: false,
      admittedBlockExtentPt: 20,
    });
  });

  it('relocates when final visible content itself crosses the region edge', () => {
    const selected = selectParagraphFragment(
      twoLineParagraph(10), { boundary: null }, splittable(twoLineBoundaries),
      19, 40, true,
      { keepLines: false, widowControl: true, authoredSpaceAfterPt: 10 },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('keeps retained trailing border extent in the final-fragment fit decision', () => {
    const selected = selectParagraphFragment(
      twoLineParagraph(5, 12), { boundary: null }, splittable(twoLineBoundaries),
      25, 40, true,
      { keepLines: false, widowControl: true, authoredSpaceAfterPt: 5 },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('keeps page-local reserve in the final-fragment fit decision', () => {
    const selected = selectParagraphFragment(
      twoLineParagraph(10), { boundary: null }, splittable(twoLineBoundaries),
      20, 40, true,
      { keepLines: false, widowControl: true, authoredSpaceAfterPt: 10 },
      () => 1,
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('re-evaluates widow removal until the fragment satisfies orphan control', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, splittable([
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ]), 20, 40, true, { keepLines: false, widowControl: true },
    );

    expect(selected).toMatchObject({
      fragment: null,
      nextCursor: { boundary: null },
      requiresFreshFlowRegion: true,
    });
  });

  it('admits only the fresh-region extent when an indivisible first line must make progress', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, splittable([
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ]), 5, 5, false, { keepLines: false, widowControl: false },
    );

    expect(selected.fragment?.advancePt).toBe(10);
    expect(selected.admittedBlockExtentPt).toBe(5);
  });

  it('selects against the reserve owned by each candidate slice', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, splittable([
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ]), 25, 40, false, { keepLines: false, widowControl: false },
      (fragment) => fragment.lines.length >= 2 ? 10 : 0,
    );

    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.additionalReservePt).toBe(0);
  });

  it('keeps a note-bearing slice out of a page whose global reserve cannot grow', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, splittable([
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ]), 40, 40, false, { keepLines: false, widowControl: false },
      (fragment) => fragment.lines.length >= 2 ? 10 : 0,
      undefined,
      (reservePt) => reservePt <= 5,
    );

    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.additionalReservePt).toBe(0);
  });

  it('keeps the explicit line break source gap when continuing onto another page', () => {
    const original = paragraph();
    const withBreak: ParagraphLayout = {
      ...original,
      lines: original.lines.map((line, index) => ({
        ...line,
        range: { start: index * 2, end: index * 2 + 1 },
      })),
    };
    const selected = selectParagraphFragment(
      withBreak, { boundary: null }, splittable([
        { segIndex: 1, charOffset: 0 },
        { segIndex: 3, charOffset: 0 },
        { segIndex: 4, charOffset: 1 },
      ]), 15, 40, false, { keepLines: false, widowControl: false },
    );
    expect(selected.fragment?.lines[0]?.range).toEqual({ start: 0, end: 1 });
    expect(selected.nextCursor?.sourceRangeStart).toBe(2);
  });

  it('carries the uniform ruby advance into the exact next source cursor', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, splittable([
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ]), 15, 40, false, { keepLines: false, widowControl: false },
      undefined,
      30,
    );

    expect(selected.nextCursor).toEqual({
      boundary: { segIndex: 0, charOffset: 1 },
      sourceRangeStart: 1,
      uniformRubyAdvancePt: 30,
    });
  });
});
