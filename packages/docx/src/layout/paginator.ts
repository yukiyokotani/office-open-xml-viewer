import {
  sectionContentStartBlockPt,
  type PageFlowSectionContext,
} from './context.js';

export type PageAdvanceReason =
  | 'overflow'
  | 'explicit-break'
  | 'section-break'
  | 'parity';

export type SectionStartType = 'continuous' | 'nextPage' | 'oddPage' | 'evenPage';

export interface PageFlowState {
  readonly pageIndex: number;
  readonly columnIndex: number;
  /** Page-absolute logical block coordinate (pt), independent of writing mode. */
  readonly cursorBlockPt: number;
  /** Logical block origin of the physical page's body content. */
  readonly pageContentStartBlockPt: number;
  /** Logical block origin shared by every column in the active section region. */
  readonly regionStartBlockPt: number;
  /** Deepest block edge reached by any completed/current column in the region. */
  readonly deepestColumnBlockPt: number;
  readonly section: PageFlowSectionContext;
}

export type PageFlowEvent =
  | Readonly<{ type: 'next-column' }>
  | Readonly<{
      type: 'next-page';
      reason: PageAdvanceReason;
      pageIndex: number;
      sectionOccurrenceId: string;
      parityBlank: boolean;
    }>
  | Readonly<{ type: 'begin-section'; section: PageFlowSectionContext }>;

export interface PageFlowTransition {
  readonly state: PageFlowState;
  readonly events: readonly PageFlowEvent[];
}

export function createPageFlowState(
  section: PageFlowSectionContext,
  overrides: Partial<Omit<PageFlowState, 'section'>> = {},
): PageFlowState {
  const contentStart = sectionContentStartBlockPt(section);
  const pageContentStartBlockPt = overrides.pageContentStartBlockPt ?? contentStart;
  const regionStartBlockPt = overrides.regionStartBlockPt ?? pageContentStartBlockPt;
  const cursorBlockPt = overrides.cursorBlockPt ?? regionStartBlockPt;
  return Object.freeze({
    pageIndex: overrides.pageIndex ?? 0,
    columnIndex: overrides.columnIndex ?? 0,
    cursorBlockPt,
    pageContentStartBlockPt,
    regionStartBlockPt,
    deepestColumnBlockPt: overrides.deepestColumnBlockPt ?? cursorBlockPt,
    section,
  });
}

function transition(
  state: PageFlowState,
  events: readonly PageFlowEvent[],
): PageFlowTransition {
  return Object.freeze({ state, events: Object.freeze([...events]) });
}

export function advanceColumnOrPage(
  state: PageFlowState,
  reason: Extract<PageAdvanceReason, 'overflow' | 'explicit-break'>,
): PageFlowTransition {
  const deepestColumnBlockPt = Math.max(
    state.deepestColumnBlockPt,
    state.cursorBlockPt,
  );
  if (state.columnIndex + 1 < state.section.columns.length) {
    return transition(Object.freeze({
      ...state,
      columnIndex: state.columnIndex + 1,
      cursorBlockPt: state.regionStartBlockPt,
      deepestColumnBlockPt,
    }), [{ type: 'next-column' }]);
  }

  const pageIndex = state.pageIndex + 1;
  return transition(createPageFlowState(state.section, { pageIndex }), [{
    type: 'next-page',
    reason,
    pageIndex,
    sectionOccurrenceId: state.section.sectionOccurrenceId,
    parityBlank: false,
  }]);
}

function matchesParity(pageIndex: number, startType: 'oddPage' | 'evenPage'): boolean {
  const isOddPhysicalPage = pageIndex % 2 === 0;
  return startType === 'oddPage' ? isOddPhysicalPage : !isOddPhysicalPage;
}

export function beginSection(
  state: PageFlowState,
  section: PageFlowSectionContext,
  startType: SectionStartType,
): PageFlowTransition {
  if (startType === 'continuous') {
    // §17.6.4: a section following newspaper columns begins below the deepest
    // column, not merely below the last column visited by source order.
    const regionTop = Math.max(state.cursorBlockPt, state.deepestColumnBlockPt);
    return transition(createPageFlowState(section, {
      pageIndex: state.pageIndex,
      pageContentStartBlockPt: state.pageContentStartBlockPt,
      cursorBlockPt: regionTop,
      regionStartBlockPt: regionTop,
      deepestColumnBlockPt: regionTop,
    }), [{ type: 'begin-section', section }]);
  }

  let pageIndex = state.pageIndex + 1;
  const events: PageFlowEvent[] = [];
  if (
    (startType === 'oddPage' || startType === 'evenPage')
    && !matchesParity(pageIndex, startType)
  ) {
    // §17.18.77: parity padding precedes the incoming section, so the blank page
    // retains the outgoing section context while the following page owns the new one.
    events.push({
      type: 'next-page',
      reason: 'parity',
      pageIndex,
      sectionOccurrenceId: state.section.sectionOccurrenceId,
      parityBlank: true,
    });
    pageIndex += 1;
  }
  events.push({
    type: 'next-page',
    reason: 'section-break',
    pageIndex,
    sectionOccurrenceId: section.sectionOccurrenceId,
    parityBlank: false,
  });
  events.push({ type: 'begin-section', section });
  return transition(createPageFlowState(section, { pageIndex }), events);
}
