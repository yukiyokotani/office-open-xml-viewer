import {
  accumulatePageSectionRegion,
  createLayoutPageAccumulator,
  type LayoutPageAccumulator,
  type LayoutPageAccumulatorInput,
  type PageSectionRegionInput,
} from './page-factory.js';
import type { PageFlowEvent, PageFlowState, PageFlowTransition } from './paginator.js';

export type CanonicalPageDraft = Readonly<{
  kind: 'content' | 'parity-blank';
  accumulator: LayoutPageAccumulator;
}>;

export type CanonicalPageDraftInput = LayoutPageAccumulatorInput & Readonly<{
  kind: CanonicalPageDraft['kind'];
  region?: PageSectionRegionInput;
}>;

export interface BodyPaginationState {
  readonly flow: PageFlowState;
  readonly pages: readonly CanonicalPageDraft[];
  readonly footnoteReservePt: number;
  readonly balanceTargetPt: number | null;
}

type NextPageEvent = Extract<PageFlowEvent, { type: 'next-page' }>;
type BeginSectionEvent = Extract<PageFlowEvent, { type: 'begin-section' }>;

export interface BodyPageTransitionFactory {
  openContentPage(
    event: NextPageEvent,
    flow: PageFlowState,
  ): Readonly<{ page: CanonicalPageDraft; flow: PageFlowState }>;
  openParityBlankPage(event: NextPageEvent): CanonicalPageDraft;
  openSamePageSectionRegion(
    page: CanonicalPageDraft,
    event: BeginSectionEvent,
    flow: PageFlowState,
  ): CanonicalPageDraft;
}

function retain(input: BodyPaginationState): BodyPaginationState {
  return Object.freeze({
    ...input,
    pages: Object.freeze([...input.pages]),
  });
}

export function createCanonicalPageDraft(input: CanonicalPageDraftInput): CanonicalPageDraft {
  const { kind, region, ...page } = input;
  let accumulator = createLayoutPageAccumulator(page);
  if (region) accumulator = accumulatePageSectionRegion(accumulator, region);
  if (kind === 'content' && accumulator.sectionRegions.length === 0) {
    throw new RangeError('A content page draft requires an initial section region');
  }
  if (kind === 'parity-blank' && accumulator.sectionRegions.length !== 0) {
    throw new RangeError('A parity blank cannot retain a section region');
  }
  return Object.freeze({ kind, accumulator });
}

export function createBodyPaginationState(
  flow: PageFlowState,
  firstPage: CanonicalPageDraft,
): BodyPaginationState {
  if (firstPage.kind !== 'content' || firstPage.accumulator.pageIndex !== flow.pageIndex) {
    throw new Error('The initial body page must be owned by the active flow');
  }
  return retain({
    flow,
    pages: [firstPage],
    footnoteReservePt: 0,
    balanceTargetPt: null,
  });
}

export function setBodyBalanceTarget(
  state: BodyPaginationState,
  balanceTargetPt: number | null,
): BodyPaginationState {
  if (balanceTargetPt !== null && (!Number.isFinite(balanceTargetPt) || balanceTargetPt < 0)) {
    throw new RangeError('A body balance target must be finite and non-negative');
  }
  return retain({ ...state, balanceTargetPt });
}

export function addPageFootnoteReserves(
  state: BodyPaginationState,
  additionalExtentsPt: readonly number[],
): BodyPaginationState {
  let footnoteReservePt = state.footnoteReservePt;
  for (const additionalPt of additionalExtentsPt) {
    if (!Number.isFinite(additionalPt) || additionalPt < 0) {
      throw new RangeError('A footnote reserve increment must be finite and non-negative');
    }
    footnoteReservePt += additionalPt;
  }
  if (!Number.isFinite(footnoteReservePt)) {
    throw new RangeError('A footnote reserve must be finite');
  }
  // Preserve each note's addition order without copying the page sequence
  // once per note in a multi-reference group.
  return footnoteReservePt === state.footnoteReservePt ? state : retain({
    ...state,
    footnoteReservePt,
  });
}

/** The reducer is the sole page-sequence owner. Acquisition can request a
 * transition, but it cannot mutate or append a canonical page draft. */
export function commitPageFlowTransition(
  state: BodyPaginationState,
  transition: PageFlowTransition,
  factory: BodyPageTransitionFactory,
): BodyPaginationState {
  const pages = [...state.pages];
  let activeFlow = transition.state;
  let openedPage = false;
  for (const event of transition.events) {
    if (event.type === 'place') throw new Error('Occurrence acceptance owns place events');
    if (event.type === 'next-column') continue;
    if (event.type === 'next-page') {
      if (event.parityBlank) {
        pages.push(factory.openParityBlankPage(event));
      } else {
        const opened = factory.openContentPage(event, transition.state);
        pages.push(opened.page);
        activeFlow = opened.flow;
      }
      openedPage = true;
      continue;
    }
    if (!openedPage) {
      const current = pages.at(-1);
      if (!current || current.kind !== 'content') {
        throw new Error('A continuous section requires an active content page');
      }
      pages[pages.length - 1] = factory.openSamePageSectionRegion(
        current,
        event,
        activeFlow,
      );
    }
  }
  const active = pages.at(-1);
  if (!active || active.kind !== 'content' || active.accumulator.pageIndex !== activeFlow.pageIndex) {
    throw new Error('A page transition must end on the active content page');
  }
  return retain({
    ...state,
    flow: activeFlow,
    pages,
    footnoteReservePt: openedPage ? 0 : state.footnoteReservePt,
    balanceTargetPt: openedPage ? null : state.balanceTargetPt,
  });
}
