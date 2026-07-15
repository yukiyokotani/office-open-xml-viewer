import { describe, expect, it } from 'vitest';
import {
  createPageFlowSectionContext,
  type PageFlowSectionContext,
} from './context.js';
import {
  advanceColumnOrPage,
  beginSection,
  createPageFlowState,
} from './paginator.js';

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

describe('immutable DOCX page-flow transitions', () => {
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
