import { describe, expect, it } from 'vitest';
import { wordGridLineAdmissionExtentPt } from './body-pagination-compatibility.js';
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
    baselinePt: index * 10 + 8, advancePt: 10,
    placements: [{
      kind: 'text', text: 'A', range: { start: index, end: index + 1 },
      origin: { xPt: 0, yPt: index * 10 + 8 },
      bounds: { xPt: 0, yPt: index * 10, widthPt: 5, heightPt: 10 },
      advancePt: 5, clusters: [], paintOps: [],
      color: { kind: 'default' },
      fontRoute: { familyList: 'serif', scope: 'native', fingerprint: 'serif' },
      fontSizePt: 10, fontWeight: 400, fontStyle: 'normal',
      direction: 'ltr', writingMode: 'horizontal-tb', decorations: [],
    }],
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
      lines: paragraph().lines.slice(0, 1).map(line => ({ ...line, placements: [] })),
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

  it.each([false, true])('admits a continuing grid line while retaining preceding cells and shading=%s', shaded => {
    const retained = paragraph();
    const selected = selectParagraphFragment(
      { ...retained,
        ...(shaded ? { shading: { color: '#EEEEEE' } } : {}),
        lines: retained.lines.map(line => ({ ...line, gridAfterLeadingPt: 2 })) },
      { boundary: null },
      splittable([
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ]),
      18, 40, true,
      { keepLines: false, widowControl: false, unreservedBodyBottomPt: 18 },
    );
    // Shading paints the complete retained cell, even across a continuation.
    expect(selected.fragment?.lines).toHaveLength(shaded ? 1 : 2);
    expect(selected.fragment?.advancePt).toBe(shaded ? 10 : 20);
    expect(selected.admissionExtentPt).toBe(shaded ? 10 : 18);
    expect(selected.nextCursor?.boundary).toEqual({ segIndex: 0, charOffset: shaded ? 1 : 2 });
  });

  it('admits a visible grid prefix even when the following break line is blank', () => {
    const retained = twoLineParagraph(0);
    const selected = selectParagraphFragment({
      ...retained,
      lines: retained.lines.map((line, index) => ({
        ...line, gridAfterLeadingPt: 2,
        placements: index === 0 ? line.placements : [],
      })),
    }, { boundary: null }, splittable(twoLineBoundaries), 8, 40, true, {
      keepLines: false, widowControl: false, unreservedBodyBottomPt: 8,
    });
    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.admissionExtentPt).toBe(8);
    expect(selected.nextCursor?.boundary).toEqual(twoLineBoundaries[0]);
  });

  it.each([
    ['candidate note band', 18, 2, false],
    ['existing reservation or shortened column', undefined, 0, false],
    ['bottom wrap exclusion', 18, 0, true],
  ] as const)('charges continuing grid cells against a %s', (_label, bodyBottomPt, reservePt, excluded) => {
    const retained = paragraph();
    const selected = selectParagraphFragment({
      ...retained,
      lines: retained.lines.map(line => ({ ...line, gridAfterLeadingPt: 2 })),
      exclusions: excluded ? [{ id: 'bottom', wrap: 'topAndBottom',
        bounds: { xPt: 0, yPt: 17, widthPt: 100, heightPt: 1 }, polygon: [] }] : [],
    }, { boundary: { segIndex: 0, charOffset: 0 } }, splittable([
      { segIndex: 0, charOffset: 1 },
      { segIndex: 0, charOffset: 2 },
      { segIndex: 0, charOffset: 3 },
    ]), 18 + reservePt, 40, true, {
      keepLines: false, widowControl: false,
      ...(bodyBottomPt === undefined ? {} : { unreservedBodyBottomPt: bodyBottomPt }),
    }, () => reservePt);
    // A discounted second cell would fit exactly, but the retained 20pt
    // prefix cannot cross a reservation even on a continued paragraph.
    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.admissionExtentPt).toBe(10);
    expect(selected.nextCursor?.boundary).toEqual({ segIndex: 0, charOffset: 1 });
  });

  it.each([[4, 8, 'page'], [17, 10, 'page'], [4, 10, 'host']] as const)('checks exclusions at the keep-next successor position for exclusion top %s, extent %s, owner %s', (exclusionTop, expectedExtent, verticalOwnership) => {
    const retained = paragraph();
    const successor = {
      ...retained,
      lines: [{ ...retained.lines[0], gridAfterLeadingPt: 2 }],
      exclusions: [{ id: 'float', wrap: 'topAndBottom' as const, verticalOwnership,
        bounds: { xPt: 0, yPt: exclusionTop, widthPt: 100, heightPt: 5 }, polygon: [] }],
    };
    // Lookahead acquired the successor at y=0; a 10pt heading places its
    // actual line at y=10. A page-owned exclusion ending at 9pt is above
    // that line; a host-owned exclusion moves down with its paragraph.
    expect(wordGridLineAdmissionExtentPt(successor, 10, 18, 0, 10)).toBe(expectedExtent);
  });

  it.each(['paragraph', 'line', 'text', 'glyph'] as const)(
    'charges the full cell for an unknown %s paint capability', owner => {
      const retained = verticalEdgeParagraph('horizontal-tb', 8, undefined, 2);
      const line = { ...retained.lines[0]!, gridAfterLeadingPt: 2 };
      const candidate = { ...retained, lines: [line] };
      expect(wordGridLineAdmissionExtentPt(candidate, 10, 18)).toBe(8);
      // New retained paint must be reviewed before it joins the plain-text
      // observation, even when an older admission implementation cannot paint it.
      Object.assign(owner === 'paragraph' ? candidate
        : owner === 'line' ? line
          : owner === 'text' ? line.placements[0]!
            : (line.placements[0]! as import('./types.js').TextPlacement).paintOps[0]!,
      { futureDecoration: {} });
      expect(wordGridLineAdmissionExtentPt(candidate, 10, 18)).toBe(10);
    },
  );

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

  it('keeps an anchor-line limit when completion has a smaller reserve than its prefix', () => {
    const retained = twoLineParagraph(0);
    const selected = selectParagraphFragment(
      retained, { boundary: null }, splittable(twoLineBoundaries),
      40, 80, true, { keepLines: false, widowControl: false, lineEndLimit: 1 },
      (fragment) => fragment.continuation?.continuesOnNext ? 46 : 16,
    );
    expect(selected.fragment).toBeNull();
    expect(selected.requiresFreshFlowRegion).toBe(true);
  });

  it('finalizes a completing continuation through the source slice metadata', () => {
    const retained = twoLineParagraph(0);
    const resumed = { ...retained, lines: retained.lines.map((line, index) => ({
      ...line, range: { start: 4 + index, end: 5 + index },
    })) };
    const selected = selectParagraphFragment(
      resumed, { boundary: { segIndex: 0, charOffset: 4 }, sourceRangeStart: 4 },
      splittable([{ segIndex: 0, charOffset: 5 }, { segIndex: 0, charOffset: 6 }]), 40, 80, true,
      { keepLines: false, widowControl: false },
      (fragment) => fragment.continuation?.continuesOnNext ? 46 : 16,
    );
    expect(selected.fragment?.continuation?.continuesFromPrevious).toBe(true);
    expect(selected.fragment?.continuation?.continuesOnNext).toBe(false);
    expect(selected.fragment?.lines.map(line => line.range)).toEqual([{ start: 4, end: 5 }, { start: 5, end: 6 }]);
    expect(selected.nextCursor).toBeNull();
    expect(selected.additionalReservePt).toBe(16);
  });
});
