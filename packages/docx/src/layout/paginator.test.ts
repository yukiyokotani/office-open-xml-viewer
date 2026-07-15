import { describe, expect, it } from 'vitest';
import {
  createPageFlowSectionContext,
  type PageFlowSectionContext,
} from './context.js';
import {
  applyAuthoredBreak,
  advanceColumnOrPage,
  beginSection,
  createPageFlowState,
  placeFlowNode,
  UnsupportedPageFlowTransitionError,
} from './paginator.js';
import type { DrawingLayout } from './types.js';

function section(
  sectionOccurrenceId: string,
  options: Readonly<{
    pageWidth?: number;
    pageHeight?: number;
    marginTop?: number;
    columns?: readonly Readonly<{ xPt: number; wPt: number }>[];
    textDirection?: string;
  }> = {},
): PageFlowSectionContext {
  return createPageFlowSectionContext({
    sectionOccurrenceId,
    geometry: {
      pageWidth: options.pageWidth ?? 612,
      pageHeight: options.pageHeight ?? 792,
      marginTop: options.marginTop ?? 72,
      marginRight: 72,
      marginBottom: 72,
      marginLeft: 72,
      headerDistance: 36,
      footerDistance: 36,
    },
    columns: options.columns ?? [{ xPt: 72, wPt: 468 }],
    textDirection: options.textDirection ?? 'lrTb',
  });
}

function drawingNode(id: string, advancePt: number): DrawingLayout {
  const bounds = { xPt: 72, yPt: 120, widthPt: 80, heightPt: advancePt };
  return {
    kind: 'drawing',
    id,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'page:0/region:0/column:0',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt,
    ordinaryFlow: true,
    commands: [],
  };
}

describe('immutable DOCX page-flow transitions', () => {
  it('places a node by advancing the logical block cursor and deepest edge', () => {
    const initial = createPageFlowState(section('section-0'), {
      cursorBlockPt: 120,
      deepestColumnBlockPt: 150,
    });

    const node = drawingNode('body/drawing/0', 48);
    const transition = placeFlowNode(initial, node);

    expect(transition.state).toMatchObject({
      cursorBlockPt: 168,
      deepestColumnBlockPt: 168,
      pageHasContent: true,
    });
    expect(transition.events).toEqual([{
      type: 'place',
      node,
      blockStartPt: 120,
      blockEndPt: 168,
    }]);
    expect(initial).toMatchObject({
      cursorBlockPt: 120,
      deepestColumnBlockPt: 150,
    });
    expect(Object.isFrozen(transition.events[0])).toBe(true);
  });

  it('rejects a negative logical block advance', () => {
    const initial = createPageFlowState(section('section-0'));

    expect(() => placeFlowNode(initial, drawingNode('body/drawing/0', -1)))
      .toThrow(RangeError);
  });

  it('advances to the next column at the region top and retains the deepest completed block', () => {
    const initial = createPageFlowState(section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    }), {
      cursorBlockPt: 240,
      regionStartBlockPt: 108,
      deepestColumnBlockPt: 260,
    });

    const transition = advanceColumnOrPage(initial, 'overflow');

    expect(transition.state).toMatchObject({
      pageIndex: 0,
      columnIndex: 1,
      cursorBlockPt: 108,
      pageContentStartBlockPt: 72,
      regionStartBlockPt: 108,
      deepestColumnBlockPt: 260,
      section: { sectionOccurrenceId: 'section-0' },
    });
    expect(transition.events).toEqual([{ type: 'next-column' }]);
    expect(initial.columnIndex).toBe(0);
    expect(Object.isFrozen(transition.state)).toBe(true);
  });

  it('advances from the last column to a fresh page in the same section', () => {
    const context = section('section-0', {
      marginTop: -54,
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    });
    const initial = createPageFlowState(context, {
      pageIndex: 2,
      columnIndex: 1,
      cursorBlockPt: 700,
      regionStartBlockPt: 108,
      deepestColumnBlockPt: 700,
    });

    const transition = advanceColumnOrPage(initial, 'overflow');

    expect(transition.state).toMatchObject({
      pageIndex: 3,
      columnIndex: 0,
      cursorBlockPt: 54,
      pageContentStartBlockPt: 54,
      regionStartBlockPt: 54,
      deepestColumnBlockPt: 54,
      section: { sectionOccurrenceId: 'section-0' },
    });
    expect(transition.events).toEqual([{
      type: 'next-page',
      reason: 'overflow',
      pageIndex: 3,
      sectionOccurrenceId: 'section-0',
      parityBlank: false,
    }]);
  });

  it('honors an authored column break by advancing to the next column', () => {
    const initial = createPageFlowState(section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    }), {
      cursorBlockPt: 312,
      regionStartBlockPt: 96,
      deepestColumnBlockPt: 312,
    });

    const transition = applyAuthoredBreak(initial, 'column');

    expect(transition.state).toMatchObject({
      pageIndex: 0,
      columnIndex: 1,
      cursorBlockPt: 96,
      deepestColumnBlockPt: 312,
    });
    expect(transition.events).toEqual([{ type: 'next-column' }]);
  });

  it('honors an authored column break in the last column by opening a page', () => {
    const initial = createPageFlowState(section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    }), {
      pageIndex: 4,
      columnIndex: 1,
      cursorBlockPt: 312,
      deepestColumnBlockPt: 312,
    });

    const transition = applyAuthoredBreak(initial, 'column');

    expect(transition.state).toMatchObject({
      pageIndex: 5,
      columnIndex: 0,
      cursorBlockPt: 72,
    });
    expect(transition.events).toEqual([{
      type: 'next-page',
      reason: 'explicit-break',
      pageIndex: 5,
      sectionOccurrenceId: 'section-0',
      parityBlank: false,
    }]);
  });

  it('honors an authored page break independently of the active column', () => {
    const initial = createPageFlowState(section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    }), {
      pageIndex: 2,
      columnIndex: 0,
      cursorBlockPt: 312,
    });

    const transition = applyAuthoredBreak(initial, 'page');

    expect(transition.state).toMatchObject({
      pageIndex: 3,
      columnIndex: 0,
      cursorBlockPt: 72,
    });
    expect(transition.events).toEqual([{
      type: 'next-page',
      reason: 'explicit-break',
      pageIndex: 3,
      sectionOccurrenceId: 'section-0',
      parityBlank: false,
    }]);
  });

  it('records pageBreakBefore as its own authored page-advance reason', () => {
    const initial = createPageFlowState(section('section-0'), {
      pageIndex: 6,
      cursorBlockPt: 312,
    });

    const transition = applyAuthoredBreak(initial, 'pageBreakBefore');

    expect(transition.state.pageIndex).toBe(7);
    expect(transition.events).toEqual([{
      type: 'next-page',
      reason: 'page-break-before',
      pageIndex: 7,
      sectionOccurrenceId: 'section-0',
      parityBlank: false,
    }]);
  });

  it('does not advance pageBreakBefore when its paragraph is already at page start', () => {
    const initial = createPageFlowState(section('section-0'), { pageIndex: 6 });

    const transition = applyAuthoredBreak(initial, 'pageBreakBefore');

    expect(transition.state).toBe(initial);
    expect(transition.events).toEqual([]);
  });

  it('keeps pageBreakBefore idempotent after prior-page content opened a fresh page', () => {
    const placed = placeFlowNode(
      createPageFlowState(section('section-0'), { pageIndex: 6 }),
      drawingNode('body/drawing/0', 24),
    ).state;
    const freshPage = applyAuthoredBreak(placed, 'page').state;

    const transition = applyAuthoredBreak(freshPage, 'pageBreakBefore');

    expect(freshPage.pageHasContent).toBe(false);
    expect(transition.state).toBe(freshPage);
    expect(transition.events).toEqual([]);
  });

  it('advances pageBreakBefore from a later column even when its cursor is at the region start', () => {
    const initial = createPageFlowState(section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    }));
    const withPriorColumnContent = advanceColumnOrPage(
      placeFlowNode(initial, drawingNode('body/drawing/0', 24)).state,
      'overflow',
    ).state;

    const transition = applyAuthoredBreak(withPriorColumnContent, 'pageBreakBefore');

    expect(withPriorColumnContent).toMatchObject({
      columnIndex: 1,
      cursorBlockPt: 72,
      pageHasContent: true,
    });
    expect(transition.state).toMatchObject({ pageIndex: 1, columnIndex: 0 });
    expect(transition.events[0]).toMatchObject({
      type: 'next-page',
      reason: 'page-break-before',
    });
  });

  it('ignores lastRenderedPageBreak because it is cached layout, not an authored break', () => {
    const initial = createPageFlowState(section('section-0'), {
      pageIndex: 6,
      cursorBlockPt: 312,
    });

    const transition = applyAuthoredBreak(initial, 'lastRenderedPageBreak');

    expect(transition.state).toBe(initial);
    expect(transition.events).toEqual([]);
  });

  it('starts a continuous section below both the live cursor and every completed column', () => {
    const outgoing = section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    });
    const incoming = section('section-1', {
      columns: [{ xPt: 72, wPt: 468 }],
    });
    const initial = createPageFlowState(outgoing, {
      cursorBlockPt: 310,
      regionStartBlockPt: 120,
      deepestColumnBlockPt: 540,
    });

    const transition = beginSection(initial, incoming, 'continuous');

    expect(transition.state).toMatchObject({
      pageIndex: 0,
      columnIndex: 0,
      cursorBlockPt: 540,
      pageContentStartBlockPt: 72,
      regionStartBlockPt: 540,
      deepestColumnBlockPt: 540,
      section: { sectionOccurrenceId: 'section-1' },
    });
    expect(transition.events).toEqual([{
      type: 'begin-section',
      section: incoming,
    }]);
  });

  it('starts a continuous section on the next page when the current page owns a footnote reference', () => {
    const outgoing = section('section-0');
    const incoming = section('section-1');
    const initial = createPageFlowState(outgoing, {
      pageIndex: 2,
      cursorBlockPt: 360,
      deepestColumnBlockPt: 360,
    });

    const transition = beginSection(initial, incoming, 'continuous', {
      hasFootnoteReferenceOnCurrentPage: true,
    });

    expect(transition.state).toMatchObject({
      pageIndex: 3,
      columnIndex: 0,
      cursorBlockPt: 72,
      section: { sectionOccurrenceId: 'section-1' },
    });
    expect(transition.events).toEqual([
      {
        type: 'next-page',
        reason: 'section-break',
        pageIndex: 3,
        sectionOccurrenceId: 'section-1',
        parityBlank: false,
      },
      { type: 'begin-section', section: incoming },
    ]);
  });

  it('starts a next-column section in the following column on the same page', () => {
    const outgoing = section('section-0', {
      columns: [
        { xPt: 72, wPt: 142 },
        { xPt: 235, wPt: 142 },
        { xPt: 398, wPt: 142 },
      ],
    });
    const incoming = section('section-1', {
      columns: [
        { xPt: 72, wPt: 142 },
        { xPt: 235, wPt: 142 },
        { xPt: 398, wPt: 142 },
      ],
    });
    const initial = createPageFlowState(outgoing, {
      pageIndex: 2,
      columnIndex: 0,
      cursorBlockPt: 360,
      deepestColumnBlockPt: 400,
    });

    const transition = beginSection(initial, incoming, 'nextColumn');

    expect(transition.state).toMatchObject({
      pageIndex: 2,
      columnIndex: 1,
      cursorBlockPt: 72,
      regionStartBlockPt: 72,
      deepestColumnBlockPt: 400,
      section: { sectionOccurrenceId: 'section-1' },
    });
    expect(transition.events).toEqual([
      { type: 'next-column' },
      { type: 'begin-section', section: incoming },
    ]);
  });

  it('rejects nextColumn when the outgoing column has no same-page successor', () => {
    const outgoing = section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    });
    const incoming = section('section-1', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    });
    const initial = createPageFlowState(outgoing, {
      pageIndex: 2,
      columnIndex: 1,
      cursorBlockPt: 360,
      deepestColumnBlockPt: 400,
    });

    expect(() => beginSection(initial, incoming, 'nextColumn'))
      .toThrowError(UnsupportedPageFlowTransitionError);
    try {
      beginSection(initial, incoming, 'nextColumn');
    } catch (error) {
      expect(error).toMatchObject({
        code: 'NEXT_COLUMN_DESTINATION_UNAVAILABLE',
        outgoingColumnIndex: 1,
        outgoingColumnCount: 2,
        incomingColumnCount: 2,
      });
    }
  });

  it('rejects nextColumn when the incoming section lacks the same-page successor index', () => {
    const outgoing = section('section-0', {
      columns: [
        { xPt: 72, wPt: 142 },
        { xPt: 235, wPt: 142 },
        { xPt: 398, wPt: 142 },
      ],
    });
    const incoming = section('section-1');
    const initial = createPageFlowState(outgoing, {
      pageIndex: 2,
      columnIndex: 0,
      cursorBlockPt: 360,
      deepestColumnBlockPt: 400,
    });

    expect(() => beginSection(initial, incoming, 'nextColumn'))
      .toThrowError(UnsupportedPageFlowTransitionError);
    try {
      beginSection(initial, incoming, 'nextColumn');
    } catch (error) {
      expect(error).toMatchObject({
        code: 'NEXT_COLUMN_DESTINATION_UNAVAILABLE',
        outgoingColumnIndex: 0,
        outgoingColumnCount: 3,
        incomingColumnCount: 1,
      });
    }
  });

  it('rejects invalid page, column, and cursor state at construction', () => {
    const twoColumns = section('section-0', {
      columns: [{ xPt: 72, wPt: 224 }, { xPt: 316, wPt: 224 }],
    });

    expect(() => createPageFlowState(twoColumns, { pageIndex: -1 })).toThrow(RangeError);
    expect(() => createPageFlowState(twoColumns, { pageIndex: 1.5 })).toThrow(RangeError);
    expect(() => createPageFlowState(twoColumns, { columnIndex: -1 })).toThrow(RangeError);
    expect(() => createPageFlowState(twoColumns, { columnIndex: 2 })).toThrow(RangeError);
    expect(() => createPageFlowState(twoColumns, { cursorBlockPt: Number.NaN })).toThrow(RangeError);
    expect(() => createPageFlowState(twoColumns, {
      regionStartBlockPt: 100,
      cursorBlockPt: 99,
    })).toThrow(RangeError);
    expect(() => createPageFlowState(twoColumns, {
      cursorBlockPt: 100,
      deepestColumnBlockPt: 99,
    })).toThrow(RangeError);
  });

  it('opens a fresh page for a next-page section', () => {
    const initial = createPageFlowState(section('section-0'), {
      pageIndex: 3,
      cursorBlockPt: 420,
      deepestColumnBlockPt: 420,
    });
    const incoming = section('section-1', { marginTop: 90 });

    const transition = beginSection(initial, incoming, 'nextPage');

    expect(transition.state).toMatchObject({
      pageIndex: 4,
      columnIndex: 0,
      cursorBlockPt: 90,
      pageContentStartBlockPt: 90,
      regionStartBlockPt: 90,
      deepestColumnBlockPt: 90,
      section: { sectionOccurrenceId: 'section-1' },
    });
    expect(transition.events).toEqual([
      {
        type: 'next-page',
        reason: 'section-break',
        pageIndex: 4,
        sectionOccurrenceId: 'section-1',
        parityBlank: false,
      },
      { type: 'begin-section', section: incoming },
    ]);
  });

  it.each([
    ['oddPage', 0, 2],
    ['evenPage', 1, 3],
  ] as const)(
    'keeps a parity-padding page in the outgoing section for %s',
    (startType, currentPageIndex, incomingPageIndex) => {
      const outgoing = section('section-0');
      const incoming = section('section-1');
      const initial = createPageFlowState(outgoing, { pageIndex: currentPageIndex });

      const transition = beginSection(initial, incoming, startType);

      expect(transition.state.pageIndex).toBe(incomingPageIndex);
      expect(transition.events).toEqual([
        {
          type: 'next-page',
          reason: 'parity',
          pageIndex: incomingPageIndex - 1,
          sectionOccurrenceId: 'section-0',
          parityBlank: true,
        },
        {
          type: 'next-page',
          reason: 'section-break',
          pageIndex: incomingPageIndex,
          sectionOccurrenceId: 'section-1',
          parityBlank: false,
        },
        { type: 'begin-section', section: incoming },
      ]);
    },
  );

  it('switches mixed section geometry and direction as one context', () => {
    const initial = createPageFlowState(section('section-0'));
    const incoming = section('section-1', {
      pageWidth: 792,
      pageHeight: 612,
      marginTop: 48,
      textDirection: 'tbRl',
      columns: [{ xPt: 48, wPt: 516 }],
    });

    const transition = beginSection(initial, incoming, 'nextPage');

    expect(transition.state.section).toBe(incoming);
    expect(transition.state.section).toMatchObject({
      sectionOccurrenceId: 'section-1',
      geometry: { pageWidth: 792, pageHeight: 612, marginTop: 48 },
      textDirection: 'tbRl',
      columns: [{ xPt: 48, wPt: 516 }],
    });
    expect(transition.state.cursorBlockPt).toBe(48);
  });
});
