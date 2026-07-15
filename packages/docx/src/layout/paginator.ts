import {
  sectionContentStartBlockPt,
  type PageFlowSectionContext,
} from './context.js';
import type { PaintNode } from './types.js';

export type PageAdvanceReason =
  | 'overflow'
  | 'explicit-break'
  | 'page-break-before'
  | 'section-break'
  | 'parity';

export type SectionStartType =
  | 'continuous'
  | 'nextColumn'
  | 'nextPage'
  | 'oddPage'
  | 'evenPage';

export type AuthoredBreak =
  | 'column'
  | 'page'
  | 'pageBreakBefore'
  | 'lastRenderedPageBreak';

export class UnsupportedPageFlowTransitionError extends Error {
  readonly code = 'NEXT_COLUMN_DESTINATION_UNAVAILABLE' as const;

  constructor(
    readonly outgoingColumnIndex: number,
    readonly outgoingColumnCount: number,
    readonly incomingColumnCount: number,
  ) {
    super(
      'nextColumn requires a following column on the current page, '
      + `but column ${outgoingColumnIndex + 1} is unavailable `
      + `(outgoing columns: ${outgoingColumnCount}, incoming columns: ${incomingColumnCount})`,
    );
    this.name = 'UnsupportedPageFlowTransitionError';
  }
}

export interface SectionBoundaryOptions {
  readonly hasFootnoteReferenceOnCurrentPage?: boolean;
}

export interface PageFlowState {
  readonly pageIndex: number;
  readonly columnIndex: number;
  /** Whether this physical page already owns placed body content. */
  readonly pageHasContent: boolean;
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
  | Readonly<{
      type: 'place';
      node: PaintNode;
      blockStartPt: number;
      blockEndPt: number;
    }>
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
  const deepestColumnBlockPt = overrides.deepestColumnBlockPt ?? cursorBlockPt;
  const pageIndex = overrides.pageIndex ?? 0;
  const columnIndex = overrides.columnIndex ?? 0;
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new RangeError('Page index must be a non-negative integer');
  }
  if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= section.columns.length) {
    throw new RangeError('Column index must identify a column in the active section');
  }
  if (![pageContentStartBlockPt, regionStartBlockPt, cursorBlockPt, deepestColumnBlockPt]
    .every(Number.isFinite)) {
    throw new RangeError('Page-flow cursors must be finite');
  }
  if (
    pageContentStartBlockPt > regionStartBlockPt
    || regionStartBlockPt > cursorBlockPt
    || cursorBlockPt > deepestColumnBlockPt
  ) {
    throw new RangeError(
      'Page-flow cursors must be ordered page start <= region start <= cursor <= deepest edge',
    );
  }
  return Object.freeze({
    pageIndex,
    columnIndex,
    pageHasContent: overrides.pageHasContent ?? false,
    cursorBlockPt,
    pageContentStartBlockPt,
    regionStartBlockPt,
    deepestColumnBlockPt,
    section,
  });
}

function transition(
  state: PageFlowState,
  events: readonly PageFlowEvent[],
): PageFlowTransition {
  return Object.freeze({
    state,
    events: Object.freeze(events.map((event) => Object.freeze({ ...event }))),
  });
}

export function placeFlowNode(
  state: PageFlowState,
  node: PaintNode,
): PageFlowTransition {
  if (!Number.isFinite(node.advancePt) || node.advancePt < 0) {
    throw new RangeError('A flow node block advance must be a finite non-negative value');
  }
  const blockStartPt = state.cursorBlockPt;
  const blockEndPt = blockStartPt + node.advancePt;
  return transition(Object.freeze({
    ...state,
    pageHasContent: true,
    cursorBlockPt: blockEndPt,
    deepestColumnBlockPt: Math.max(state.deepestColumnBlockPt, blockEndPt),
  }), [{
    type: 'place',
    node,
    blockStartPt,
    blockEndPt,
  }]);
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

function advanceToPage(
  state: PageFlowState,
  section: PageFlowSectionContext,
  reason: Extract<PageAdvanceReason, 'explicit-break' | 'page-break-before' | 'section-break'>,
): PageFlowTransition {
  const pageIndex = state.pageIndex + 1;
  return transition(createPageFlowState(section, { pageIndex }), [{
    type: 'next-page',
    reason,
    pageIndex,
    sectionOccurrenceId: section.sectionOccurrenceId,
    parityBlank: false,
  }]);
}

export function applyAuthoredBreak(
  state: PageFlowState,
  authoredBreak: AuthoredBreak,
): PageFlowTransition {
  if (authoredBreak === 'lastRenderedPageBreak') {
    // lastRenderedPageBreak is a cached result from a previous layout producer,
    // not document intent. Mixing it with fresh pagination double-applies breaks.
    return transition(state, []);
  }
  if (authoredBreak === 'column') {
    return advanceColumnOrPage(state, 'explicit-break');
  }
  if (
    authoredBreak === 'pageBreakBefore'
    && !state.pageHasContent
    && state.columnIndex === 0
    && state.cursorBlockPt === state.pageContentStartBlockPt
  ) {
    // §17.3.1.23 requires the paragraph to begin on a new page. A paragraph
    // already at the start of an otherwise empty page satisfies that condition.
    return transition(state, []);
  }
  return advanceToPage(
    state,
    state.section,
    authoredBreak === 'pageBreakBefore' ? 'page-break-before' : 'explicit-break',
  );
}

function matchesParity(pageIndex: number, startType: 'oddPage' | 'evenPage'): boolean {
  const isOddPhysicalPage = pageIndex % 2 === 0;
  return startType === 'oddPage' ? isOddPhysicalPage : !isOddPhysicalPage;
}

export function beginSection(
  state: PageFlowState,
  section: PageFlowSectionContext,
  startType: SectionStartType,
  options: SectionBoundaryOptions = {},
): PageFlowTransition {
  if (startType === 'continuous' && !options.hasFootnoteReferenceOnCurrentPage) {
    // §17.6.4: a section following newspaper columns begins below the deepest
    // column, not merely below the last column visited by source order.
    const regionTop = Math.max(state.cursorBlockPt, state.deepestColumnBlockPt);
    return transition(createPageFlowState(section, {
      pageIndex: state.pageIndex,
      pageContentStartBlockPt: state.pageContentStartBlockPt,
      cursorBlockPt: regionTop,
      regionStartBlockPt: regionTop,
      deepestColumnBlockPt: regionTop,
      pageHasContent: state.pageHasContent,
    }), [{ type: 'begin-section', section }]);
  }

  if (startType === 'nextColumn') {
    const followingColumnIndex = state.columnIndex + 1;
    if (
      followingColumnIndex < state.section.columns.length
      && followingColumnIndex < section.columns.length
    ) {
      return transition(Object.freeze({
        ...state,
        columnIndex: followingColumnIndex,
        cursorBlockPt: state.pageContentStartBlockPt,
        regionStartBlockPt: state.pageContentStartBlockPt,
        deepestColumnBlockPt: Math.max(
          state.deepestColumnBlockPt,
          state.cursorBlockPt,
        ),
        section,
      }), [
        { type: 'next-column' },
        { type: 'begin-section', section },
      ]);
    }

    // §17.18.77 only defines nextColumn when a following column exists on this
    // page. No normative rule selects another page when that destination is
    // absent, so the caller must surface the unsupported transition explicitly.
    throw new UnsupportedPageFlowTransitionError(
      state.columnIndex,
      state.section.columns.length,
      section.columns.length,
    );
  }

  if (startType === 'continuous') {
    // §17.18.77 requires the continuous section to begin on the following page
    // when a footnote reference on this page would otherwise cross the boundary.
    const nextPage = advanceToPage(state, section, 'section-break');
    return transition(nextPage.state, [
      ...nextPage.events,
      { type: 'begin-section', section },
    ]);
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
