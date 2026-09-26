import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type { PageBorders } from '../types.js';
import { createPageFlowSectionContext } from './context.js';
import { finalizeLayoutPage } from './page-factory.js';
import {
  advanceColumnOrPage,
  beginSection,
  createPageFlowState,
} from './paginator.js';
import {
  addPageFootnoteReserves,
  commitPageFlowTransition,
  createBodyPaginationState,
  createCanonicalPageDraft,
  setBodyBalanceTarget,
} from './body-pagination.js';

const section: SectionLayoutContext = {
  geometry: {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5,
  },
  columns: [{ xPt: 10, wPt: 180 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'lrTb', verticalAlignment: 'top',
};

const draft = (pageIndex: number) => createCanonicalPageDraft({
  kind: 'content',
  pageIndex,
  physicalPage: {
    widthPt: 200, heightPt: 100, contentTopPt: 10, contentBottomPt: 90,
  },
  sectionOccurrenceId: 'section:0',
  section,
  region: {
    id: `region:${pageIndex}`,
    sectionOccurrenceId: 'section:0',
    section,
    writingMode: 'horizontal-tb',
    blockStartPt: 10,
    blockEndPt: 90,
    columns: [{ inlineStartPt: 10, inlineExtentPt: 180 }],
  },
});

describe('immutable canonical page transitions', () => {
  it.each([
    ['allPages', true],
    ['firstPage', true],
    ['notFirstPage', false],
  ] as const)(
    'materializes parser-owned %s page borders for the first section-owned page',
    (display, visible) => {
      const pageBorders: PageBorders = {
        offsetFrom: 'page', display, zOrder: 'front',
        top: { style: 'single', width: 1, space: 12 },
      };
      const pageDraft = createCanonicalPageDraft({
        kind: 'content',
        pageIndex: 0,
        physicalPage: {
          widthPt: 200, heightPt: 100, contentTopPt: 10, contentBottomPt: 90,
        },
        sectionOccurrenceId: 'section:0',
        section,
        region: {
          id: 'region:0', sectionOccurrenceId: 'section:0', section,
          pageBorders,
          writingMode: 'horizontal-tb', blockStartPt: 10, blockEndPt: 90,
          columns: [{ inlineStartPt: 10, inlineExtentPt: 180 }],
        },
      });

      const page = finalizeLayoutPage(pageDraft.accumulator, {
        displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0',
      }, true);

      expect(page.pageBorder !== null).toBe(visible);
      if (visible) {
        expect(page.pageBorder).toMatchObject({
          zOrder: 'front',
          segments: [{
            edge: 'top',
            from: { xPt: 0, yPt: 12 },
            to: { xPt: 200, yPt: 12 },
            widthPt: 1,
          }],
        });
      }
    },
  );

  it('requires document finalization to supply page-border section ownership', () => {
    const pageDraft = createCanonicalPageDraft({
      kind: 'content',
      pageIndex: 4,
      physicalPage: {
        widthPt: 200, heightPt: 100, contentTopPt: 10, contentBottomPt: 90,
      },
      sectionOccurrenceId: 'section:later',
      section,
      region: {
        id: 'region:later',
        sectionOccurrenceId: 'section:later',
        section,
        pageBorders: {
          offsetFrom: 'page',
          display: 'firstPage',
          zOrder: 'front',
          top: { style: 'single', width: 1, space: 12 },
        },
        writingMode: 'horizontal-tb',
        blockStartPt: 10,
        blockEndPt: 90,
        columns: [{ inlineStartPt: 10, inlineExtentPt: 180 }],
      },
    });

    expect(() => finalizeLayoutPage(pageDraft.accumulator, {
      displayNumber: 5, format: 'decimal', sectionOccurrenceId: 'section:later',
    })).toThrow(/section-owned page identity/);
  });

  it('retains per-note arithmetic order and rejects invalid reserve groups atomically', () => {
    const flowSection = createPageFlowSectionContext({
      sectionOccurrenceId: 'section:0',
      geometry: section.geometry,
      columns: section.columns,
      textDirection: section.textDirection,
    });
    const initial = createBodyPaginationState(createPageFlowState(flowSection), draft(0));
    const first = addPageFootnoteReserves(initial, [0.1]);
    const next = addPageFootnoteReserves(first, [0.2, 0.3]);
    expect(next.footnoteReservePt).toBe(0.1 + 0.2 + 0.3);
    expect(first.footnoteReservePt).toBe(0.1);
    expect(initial.footnoteReservePt).toBe(0);
    expect(next.pages).toEqual(first.pages);
    expect(Object.isFrozen(next)).toBe(true);
    expect(Object.isFrozen(next.pages)).toBe(true);
    expect(addPageFootnoteReserves(first, [])).toBe(first);
    expect(addPageFootnoteReserves(first, [0, 0])).toBe(first);
    for (const invalid of [-1, NaN, Infinity]) {
      expect(() => addPageFootnoteReserves(first, [0.2, invalid])).toThrow(RangeError);
      expect(first.footnoteReservePt).toBe(0.1);
    }
    expect(() => addPageFootnoteReserves(initial, [Number.MAX_VALUE, Number.MAX_VALUE]))
      .toThrow(/must be finite/);
  });

  it('opens a new draft without mutating the prior state or page', () => {
    const flowSection = createPageFlowSectionContext({
      sectionOccurrenceId: 'section:0',
      geometry: section.geometry,
      columns: section.columns,
      textDirection: section.textDirection,
    });
    const originalPage = draft(0);
    const original = setBodyBalanceTarget(
      addPageFootnoteReserves(
        createBodyPaginationState(createPageFlowState(flowSection), originalPage),
        [12],
      ),
      48,
    );
    const transition = advanceColumnOrPage(original.flow, 'overflow');

    const next = commitPageFlowTransition(original, transition, {
      openContentPage: (event) => ({ page: draft(event.pageIndex), flow: transition.state }),
      openParityBlankPage: () => { throw new Error('unused'); },
      openSamePageSectionRegion: () => { throw new Error('unused'); },
    });

    expect(original.pages).toEqual([originalPage]);
    expect(original.footnoteReservePt).toBe(12);
    expect(original.balanceTargetPt).toBe(48);
    expect(next.pages.map((page) => page.accumulator.pageIndex)).toEqual([0, 1]);
    expect(next.flow.pageIndex).toBe(1);
    expect(next.footnoteReservePt).toBe(0);
    expect(next.balanceTargetPt).toBeNull();
    expect(Object.isFrozen(next)).toBe(true);
    expect(Object.isFrozen(next.pages)).toBe(true);
    expect(next).not.toBe(original);
  });

  it('dispatches a same-page-column section event without treating it as a block transition', () => {
    const columns = [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }];
    const outgoingSection = { ...section, columns };
    const incomingSection = { ...section, columns };
    const outgoingFlowSection = createPageFlowSectionContext({
      sectionOccurrenceId: 'section:outgoing',
      geometry: outgoingSection.geometry,
      columns,
      textDirection: outgoingSection.textDirection,
    });
    const incomingFlowSection = createPageFlowSectionContext({
      sectionOccurrenceId: 'section:incoming',
      geometry: incomingSection.geometry,
      columns,
      textDirection: incomingSection.textDirection,
    });
    const page = createCanonicalPageDraft({
      kind: 'content',
      pageIndex: 0,
      physicalPage: {
        widthPt: 200, heightPt: 100, contentTopPt: 10, contentBottomPt: 90,
      },
      sectionOccurrenceId: 'section:outgoing',
      section: outgoingSection,
      region: {
        id: 'region:outgoing',
        sectionOccurrenceId: 'section:outgoing',
        section: outgoingSection,
        writingMode: 'horizontal-tb',
        blockStartPt: 10,
        blockEndPt: 90,
        columns: [
          { inlineStartPt: 10, inlineExtentPt: 80 },
          { inlineStartPt: 110, inlineExtentPt: 80 },
        ],
      },
    });
    const flow = createPageFlowState(outgoingFlowSection, {
      columnIndex: 0,
      cursorBlockPt: 40,
      deepestColumnBlockPt: 40,
    });
    const original = createBodyPaginationState(flow, page);
    const transition = beginSection(flow, incomingFlowSection, 'nextColumn');
    let placement: string | undefined;

    const next = commitPageFlowTransition(original, transition, {
      openContentPage: () => { throw new Error('unused'); },
      openParityBlankPage: () => { throw new Error('unused'); },
      openSamePageSectionRegion: (current, event) => {
        placement = 'placement' in event ? event.placement : undefined;
        return current;
      },
    });

    expect(placement).toBe('same-page-column');
    expect(next.flow).toMatchObject({
      pageIndex: 0,
      columnIndex: 1,
      columnSubset: [1],
      section: { sectionOccurrenceId: 'section:incoming' },
    });
  });
});
