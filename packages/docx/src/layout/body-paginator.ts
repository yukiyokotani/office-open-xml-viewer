import { MAX_BODY_LAYOUT_PAGES } from './resource-budgets.js';
import type { SectionLayoutContext } from '../layout-context.js';
import type {
  BodyLayoutInput,
  BodyParagraphSourceInput,
  BodySectionLayoutInput,
} from './body-layout-input.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutSession,
  BodyTableContinuationCursor,
  CarriedHostAnchor,
  FrameCarryObservation,
  PageAnchorPrescanInput,
} from './body-layout-kernel.js';
import { NoteCapacityExceededError, noteBandKey } from './body-layout-kernel.js';
import { NoteStack } from './note-stack.js';
import { PageNoteController, PageNoteReplay } from './page-note-controller.js';
import { acquireNoteColumnPartition, destinationOwnedNote } from './note-column-acquisition.js';
import { balancePageNoteArea, flowPageNoteArea, type PageNoteSource } from './page-note-area.js';
import { createFloatLineBandPredicate, type FloatRect } from './float-wrap.js';
import {
  commitPageFlowTransition,
  createBodyPaginationState,
  createCanonicalPageDraft,
  addPageFootnoteReserves,
  setBodyFootnoteReserve,
  setBodyBalanceTarget,
  type BodyPageTransitionFactory,
  type BodyPaginationState,
  type CanonicalPageDraft,
} from './body-pagination.js';
import { assertAndDeepFreezeDocumentLayoutSteps } from './invariants.js';
import {
  bodyLayoutKernelOf,
  createFieldAcquisitionServicesView,
  createParagraphAcquisitionCacheServicesView,
} from './runtime-state.js';
import {
  accumulatePagePaintNode,
  accumulatePageSectionRegion,
  bodyFlowDomainId,
  createParityBlankLayoutPage,
  finalizePageBookmarkStarts,
  finalizeLayoutPage,
  type PageSectionRegionInput,
} from './page-factory.js';
import {
  projectBodyOccurrence,
  projectedNestedOccurrenceId,
} from './occurrence-projection.js';
import {
  advanceColumnOrPage,
  advanceToPage,
  applyAuthoredBreak,
  beginSection,
  createPageFlowState,
  placeFlowNode,
  UnsupportedPageFlowTransitionError,
} from './paginator.js';

import {
  createPageFlowSectionContext,
  physicalSectionGeometry,
  resolveSectionContextForPage,
} from './context.js';
import {
  sectionWritingMode,
  transformRect,
  uprightPhysicalExtent,
} from './coordinate-space.js';
import {
  selectParagraphFragment,
  type ParagraphFragmentCursor,
} from './paragraph-pagination.js';
import {
  anchorKeysId,
  anchorLineDeferralApplies,
  anchorLineDeferralKey,
  anchorLineDeferralsIdentity,
  createAnchorLineDeferralProof,
  pageOwnedAnchorKeysByLine,
  serializeAnchorInput,
  type AnchorLineDeferralContext,
  type AnchorLineDeferrals,
  type AnchorLineDeferralProof,
  type PageAnchorInputEvent,
} from './anchor-line-deferral.js';
import { paragraphContinuingPrefixAdvancesPt } from './paragraph.js';
import {
  createFootnotePartitioner,
  moveContinuationNotice,
  type ContinuationNoticeProvider,
  type FootnotePartitionExtent,
  type FootnoteCursor,
} from './footnote-fragmentation.js';
import { noteSeparatorOccurrence } from './reserved-note-separators.js';
import { paragraphGapAdjustment } from './paragraph-spacing.js';
import {
  endnoteIdsInRetainedSlice,
  footnoteIdsInRetainedLines,
  footnoteIdsInRetainedSlice,
} from './note-reference-ownership.js';
import { exactRetainedColumnBalanceTarget } from './column-balance-frontier.js';
import { composeCanonicalSectionFlow } from './section-flow-composition.js';
import type { BodyFlowAllocation } from './section-flow-composition.js';
import { attachTrackChangeBars } from './track-changes.js';
import {
  isFirstSectionOwnedPage,
  sectionContentFirstAppearancePageIndices,
} from './section-page-identity.js';
import {
  wordFinalParagraphAdmissionExtentPt,
  wordGridLineAdmissionExtentPt,
  wordParagraphFootnoteMayContinue,
  wordFootnoteLookaheadBoundary,
  wordEmptyKeepNextBridgesSuccessor,
  wordFlowNeutralPreBreakAnchorParagraph,
  wordPreBreakHostAnchorExtentPt,
  wordPreBreakInlineDrawingResource,
  wordPerColumnFootnoteBandsApply,
  wordContinuousSectionMayShareFootnotePage,
} from './body-pagination-compatibility.js';
import {
  wordContinuousSectionRestartDisplayNumber,
  wordTrailingEmptyMarkAdmissionAllowancePt,
} from './section-compatibility.js';
import {
  bodyOccurrenceKey,
  bodyRootFloatingTablePlacementKey,
  sourceKey,
  tableFragmentStartKey,
} from './source-key.js';
import {
  convergeHeaderFooterReserveSteps,
  headerStoryBodyReserveExtentPt,
  headerFooterOverflowReservePt,
  reservedBodyInterval,
  selectedHeaderFooterStory,
  type HeaderFooterReserve,
  type ReservedBodyInterval,
} from './header-footer-reserve.js';
import type { LayoutOptions } from './options.js';
import type {
  BodyFlowRegistryDeltaPt,
  DeepReadonly,
  DocumentLayout,
  LayoutDiagnostic,
  LayoutPage,
  LayoutRect,
  LayoutServices,
  NoteLayout,
  PaintNode,
  ParagraphLayout,
  SourceRef,
  StoryLayout,
  TableLayout,
  WritingMode,
} from './types.js';
import { createPageLayers, pageLayerNodes, type PageLayerNode } from './page-graph.js';
import {
  translateNoteLayout,
  translateStoryLayout,
} from './stories.js';
import { LayoutInvariantError } from './diagnostics.js';
import {
  ExactConvergenceError,
  convergeExactStateSteps,
} from './convergence.js';
import { paginationFieldPageContexts } from './pagination-fields.js';
import { solveExactTranslation } from './story-page-frames.js';

// Resource governance independent of compatibility thresholds: each physical
// page retains flow/paint state. Bound adversarial page generation globally.
// The full-note acquisition and paginator share one physical-page budget.

// Retain only a small recent exact-state window, as in the reserve solve.
// This is a retention bound, not a pass limit: independent section openings
// can each require an observation. The physical-page ceiling and shared hard
// acquisition ledger bound work even for cycles longer than this window.
const MAX_OPENED_RESERVE_IDENTITIES = 16;

/** An actual empty page opening supplied new reserve observations, but the
 * body could not advance in that geometry. Retry the same reserve iteration
 * with these observations, rather than opening an unbounded run of empty pages.
 * This is a convergence signal, never a speculative fit or a copied note plan. */
class OpenedPageReserveObservation extends Error {
  constructor(readonly reserves: readonly HeaderFooterReserve[]) {
    super('An empty opened page requires reserve observation before continuing');
  }
}

class FootnoteAdmissionOverflowError extends Error {
  readonly code = 'FOOTNOTE_RESERVE_EXCEEDS_FRESH_PAGE' as const;

  constructor(
    readonly reservePt: number,
    readonly admissionChargePt: number,
    readonly freshPageExtentPt: number,
  ) {
    super(
      'Body footnote admission cannot fit a fresh physical page '
      + `(reserve: ${reservePt}, charge: ${admissionChargePt}, `
      + `fresh page: ${freshPageExtentPt})`,
    );
    this.name = 'FootnoteAdmissionOverflowError';
  }
}

interface BodyBalanceTarget {
  readonly pageIndex: number;
  readonly targetPt: number;
}

type BodyBalancePlan = ReadonlyMap<string, BodyBalanceTarget>;

function nestedStoryDiagnostics(
  node: PaintNode,
  visited: WeakSet<object>,
): readonly LayoutDiagnostic[] {
  if (visited.has(node)) return [];
  visited.add(node);
  const own = node.diagnostics ?? [];
  if (node.kind === 'paragraph') {
    return [
      ...own,
      ...node.drawings.flatMap((drawing) => nestedStoryDiagnostics(drawing, visited)),
      ...node.textBoxes.flatMap((textBox) => nestedStoryDiagnostics(textBox, visited)),
    ];
  }
  if (node.kind === 'table') {
    return [
      ...own,
      ...node.rows.flatMap((row) => row.cells.flatMap((cell) =>
        cell.blocks.flatMap((block) => nestedStoryDiagnostics(block.layout, visited)))),
      ...(node.floatingTables ?? []).flatMap((placement) =>
        nestedStoryDiagnostics(placement.child, visited)),
      ...(node.resolvedFloatingTables ?? []).flatMap((placement) =>
        nestedStoryDiagnostics(placement.child, visited)),
    ];
  }
  if (node.kind === 'textbox' || node.kind === 'note') {
    return [
      ...own,
      ...node.story.diagnostics,
      ...node.story.blocks.flatMap((block) => nestedStoryDiagnostics(block, visited)),
    ];
  }
  return own;
}

/** Notes already planned for the fragment being admitted (not committed). */
type PendingFootnotes = Readonly<{
  ids: Pick<ReadonlySet<string>, 'has' | 'size'>;
  carriesNotice: boolean;
}>;

const NO_PENDING_FOOTNOTES: PendingFootnotes =
  Object.freeze({ ids: new Set<string>(), carriesNotice: false });

/** Minimum requirements of a later reference in the same body unit: at least
 * one earlier note always precedes it on the page (so it never owns the
 * separator band), and the unit's single notice is reserved separately. Its
 * own ids are unique among the paragraph's events, so no exclusion applies. */
const LATER_UNIT_REFERENCE: PendingFootnotes = Object.freeze({
  ids: Object.freeze({ has: () => false, size: 1 }),
  carriesNotice: true,
});

/**
 * Smallest fragment end at or after `lineEnd` that ECMA-376 §17.3.1.44
 * widow/orphan control accepts for a fragment of `total` remaining lines,
 * mirroring the selector's `adjustForWidowOrphan`: a lone first line is an
 * orphan and a selection that leaves exactly one line is a widow.
 */
function widowLegalLineEnd(lineEnd: number, total: number): number {
  for (let end = lineEnd; end < total; end += 1) {
    const orphan = end === 1;
    const widow = total - end === 1 && end >= 2;
    if (!orphan && !widow) return end;
  }
  return total;
}

function assertFreshPageFootnoteAdmission(
  reservePt: number,
  admissionChargePt: number,
  freshPageExtentPt: number,
): void {
  if (reservePt > 0 && admissionChargePt > freshPageExtentPt) {
    throw new FootnoteAdmissionOverflowError(
      reservePt,
      admissionChargePt,
      freshPageExtentPt,
    );
  }
}

function nestedFloatingOccurrenceIds(layout: TableLayout): ReadonlySet<string> {
  const ids = new Set<string>();
  const visit = (table: TableLayout) => {
    for (const resolved of table.resolvedFloatingTables ?? []) {
      ids.add(resolved.occurrenceId);
      visit(resolved.child);
    }
  };
  visit(layout);
  return ids;
}

function bindTableFlowRegistryDeltaToAcceptedOccurrence(
  delta: BodyFlowRegistryDeltaPt,
  layout: TableLayout,
  ownerOccurrenceId: string,
): BodyFlowRegistryDeltaPt {
  if (!delta.floats) {
    throw new Error('Accepted floating table omitted its float registry delta');
  }
  const nestedIds = nestedFloatingOccurrenceIds(layout);
  return Object.freeze({
    ...delta,
    floats: Object.freeze({
      ...delta.floats,
      entries: Object.freeze(delta.floats.entries.map((entry) => {
        const occurrenceId = nestedIds.has(entry.occurrenceId)
          ? projectedNestedOccurrenceId(ownerOccurrenceId, entry.occurrenceId)
          : layout.ordinaryFlow
            ? null
            : ownerOccurrenceId;
        if (occurrenceId === null) return entry;
        return Object.freeze({ ...entry, occurrenceId, exclusionId: occurrenceId });
      })),
    }),
  });
}

function paragraphFlowRegistryDeltaForAcceptedFragment(
  delta: BodyFlowRegistryDeltaPt,
  fragment: ParagraphLayout,
): BodyFlowRegistryDeltaPt | null {
  const acceptedAnchorOccurrenceIds = new Set(fragment.drawings.flatMap((drawing) => {
    const occurrenceId = drawing.anchorLayer?.acquisitionOccurrenceId
      ?? drawing.anchorLayer?.occurrenceId;
    return occurrenceId === undefined ? [] : [occurrenceId];
  }));
  const floatEntries = delta.floats?.entries.filter((entry) =>
    acceptedAnchorOccurrenceIds.has(entry.occurrenceId)) ?? [];
  const collisionEntries = delta.drawingCollisions?.entries.filter((entry) =>
    acceptedAnchorOccurrenceIds.has(entry.occurrenceId)) ?? [];
  if (floatEntries.length === 0 && collisionEntries.length === 0) return null;
  return Object.freeze({
    ...(delta.floats && floatEntries.length > 0 ? {
      floats: Object.freeze({
        ...delta.floats,
        entries: Object.freeze(floatEntries),
        nextParagraphId: delta.floats.baseNextParagraphId + floatEntries.length,
      }),
    } : {}),
    ...(delta.drawingCollisions && collisionEntries.length > 0 ? {
      drawingCollisions: Object.freeze({
        ...delta.drawingCollisions,
        entries: Object.freeze(collisionEntries),
      }),
    } : {}),
  });
}

function ownerMap(input: BodyLayoutInput): Map<string, BodySectionLayoutInput> {
  const owners = new Map([[input.initialSection.sectionOccurrenceId, input.initialSection]]);
  for (let entryIndex = 0; entryIndex < input.sequence.length; entryIndex += 1) {
    const entry = input.sequence[entryIndex]!;
    if (entry.kind === 'begin-section') owners.set(entry.section.sectionOccurrenceId, entry.section);
  }
  return owners;
}

function sectionContextForPage(owner: BodySectionLayoutInput, pageIndex: number) {
  return resolveSectionContextForPage(owner.context as SectionLayoutContext, owner.pageLayout, pageIndex);
}

function flowSection(owner: BodySectionLayoutInput, pageIndex: number) {
  const context = sectionContextForPage(owner, pageIndex);
  return createPageFlowSectionContext({
    sectionOccurrenceId: owner.sectionOccurrenceId,
    geometry: context.geometry,
    columns: context.columns,
    textDirection: context.textDirection,
    ...(context.sectionFlow ? { sectionFlow: context.sectionFlow } : {}),
    sectionBidi: context.sectionBidi === true,
    grid: context.grid,
  });
}

function pageRegion(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  interval: ReservedBodyInterval,
  blockStartPt = interval.blockStartPt,
  columnIndexes: readonly number[] = sectionContextForPage(owner, pageIndex)
    .columns.map((_, index) => index),
): PageSectionRegionInput {
  const context = sectionContextForPage(owner, pageIndex);
  return Object.freeze({
    id: `page:${pageIndex}:section:${encodeURIComponent(owner.sectionOccurrenceId)}`,
    sectionOccurrenceId: owner.sectionOccurrenceId,
    section: context,
    pageBorders: owner.pageBordersAuthored ? owner.pageBorders : null,
    writingMode: sectionWritingMode(context),
    blockStartPt,
    blockEndPt: interval.blockEndPt,
    columnFlowDirection: context.sectionBidi === true ? 'rtl' : 'ltr',
    columnIndexes: Object.freeze([...columnIndexes]),
    columns: Object.freeze(columnIndexes.map((columnIndex) => {
      const column = context.columns[columnIndex];
      if (!column) throw new Error('Missing authored section column');
      return Object.freeze({
        inlineStartPt: column.xPt,
        inlineExtentPt: column.wPt,
      });
    })),
  });
}

function physicalPage(
  section: DeepReadonly<SectionLayoutContext>,
  interval: ReservedBodyInterval,
) {
  const writingMode = sectionWritingMode(section);
  const extent = uprightPhysicalExtent({
    widthPt: section.geometry.pageWidth,
    heightPt: section.geometry.pageHeight,
  }, writingMode);
  if (writingMode !== 'horizontal-tb') {
    return Object.freeze({
      ...extent,
      contentTopPt: 0,
      contentBottomPt: extent.heightPt,
    });
  }
  return Object.freeze({
    ...extent,
    contentTopPt: interval.blockStartPt,
    contentBottomPt: interval.blockEndPt,
  });
}

function pageBodyInterval(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  reserve: HeaderFooterReserve,
): ReservedBodyInterval {
  return reservedBodyInterval(sectionContextForPage(owner, pageIndex).geometry, reserve);
}

function openDraft(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  interval: ReservedBodyInterval,
): CanonicalPageDraft {
  const context = sectionContextForPage(owner, pageIndex);
  return createCanonicalPageDraft({
    kind: 'content', pageIndex,
    physicalPage: physicalPage(context, interval),
    sectionOccurrenceId: owner.sectionOccurrenceId,
    section: context,
    region: pageRegion(owner, pageIndex, interval),
  });
}

function activeRegion(state: BodyPaginationState): PageSectionRegionInput {
  const page = state.pages.at(-1);
  const region = page?.accumulator.sectionRegions.at(-1);
  if (!page || page.kind !== 'content' || !region) throw new Error('Missing active body region');
  return region;
}

function activeBlockEndPt(state: BodyPaginationState): number {
  const region = activeRegion(state);
  const columnIndexes = region.columnIndexes
    ?? region.section.columns.map((_, index) => index);
  const populationOrder = region.columnFlowDirection === 'rtl'
    ? [...columnIndexes].reverse()
    : [...columnIndexes];
  const isLastColumn = populationOrder.at(-1) === state.flow.columnIndex;
  return state.balanceTargetPt === null || isLastColumn
    ? region.blockEndPt
    : Math.min(region.blockEndPt, region.blockStartPt + state.balanceTargetPt);
}

function acquisitionLocation(state: BodyPaginationState): BodyAcquisitionLocation {
  const region = activeRegion(state);
  const columnIndexes = region.columnIndexes
    ?? region.section.columns.map((_, index) => index);
  const column = region.columns[columnIndexes.indexOf(state.flow.columnIndex)];
  if (!column) throw new Error('Missing active body column');
  return Object.freeze({
    pageIndex: state.flow.pageIndex,
    columnIndex: state.flow.columnIndex,
    flowDomainId: bodyFlowDomainId(state.flow.pageIndex, region.id, state.flow.columnIndex),
    section: region.section,
    cursorPt: Object.freeze({ xPt: column.inlineStartPt, yPt: state.flow.cursorBlockPt }),
    availableBounds: Object.freeze({
      xPt: column.inlineStartPt,
      yPt: state.flow.cursorBlockPt,
      widthPt: column.inlineExtentPt,
      heightPt: Math.max(
        0,
        activeBlockEndPt(state) - state.footnoteReservePt - state.flow.cursorBlockPt,
      ),
    }),
  });
}

function transitionFactory(
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
  reserveForOpening: (owner: BodySectionLayoutInput, pageIndex: number, parityBlank?: boolean, automaticOpening?: boolean) => HeaderFooterReserve,
): BodyPageTransitionFactory {
  const owner = (id: string) => {
    const result = owners.get(id);
    if (!result) throw new Error(`Unknown body section ${id}`);
    return result;
  };
  return {
    openContentPage(event) {
      const nextOwner = owner(event.sectionOccurrenceId);
      const reserve = reserveForOpening(nextOwner, event.pageIndex, false, event.reason === 'overflow');
      const interval = pageBodyInterval(nextOwner, event.pageIndex, reserve);
      const flow = createPageFlowState(flowSection(nextOwner, event.pageIndex), {
        pageIndex: event.pageIndex,
        pageContentStartBlockPt: interval.blockStartPt,
        pageContentEndBlockPt: interval.blockEndPt,
      });
      return { page: openDraft(nextOwner, event.pageIndex, interval), flow };
    },
    openParityBlankPage(event) {
      const blankOwner = owner(event.sectionOccurrenceId);
      const context = sectionContextForPage(blankOwner, event.pageIndex);
      const interval = pageBodyInterval(
        blankOwner,
        event.pageIndex,
        reserveForOpening(blankOwner, event.pageIndex, true),
      );
      return createCanonicalPageDraft({
        kind: 'parity-blank', pageIndex: event.pageIndex,
        physicalPage: physicalPage(context, interval),
        sectionOccurrenceId: blankOwner.sectionOccurrenceId,
        section: context,
        pageBorders: blankOwner.pageBordersAuthored ? blankOwner.pageBorders : null,
      });
    },
    openSamePageSectionRegion(page, event, flow) {
      const nextOwner = owner(event.section.sectionOccurrenceId);
      const priorRegions = page.accumulator.sectionRegions;
      const prior = priorRegions.at(-1);
      if (!prior || !('placement' in event)) {
        throw new Error('A same-page section requires explicit retained placement');
      }
      const pageInterval = Object.freeze({
        blockStartPt: page.accumulator.sectionRegions[0]!.blockStartPt,
        blockEndPt: prior.blockEndPt,
      });
      const constrainedPrior = event.placement === 'same-page-block'
        ? Object.freeze({ ...prior, blockEndPt: flow.regionStartBlockPt })
        : (() => {
            const outgoingColumnIndexes = event.outgoingColumnSubset;
            if (!outgoingColumnIndexes || outgoingColumnIndexes.length === 0) {
              throw new Error('A same-page-column transition requires outgoing column ownership');
            }
            return Object.freeze({
              ...prior,
              columnIndexes: Object.freeze([...outgoingColumnIndexes]),
              columns: Object.freeze(outgoingColumnIndexes.map((columnIndex) => {
                const column = prior.section.columns[columnIndex];
                if (!column) throw new Error('Missing outgoing authored column');
                return Object.freeze({
                  inlineStartPt: column.xPt,
                  inlineExtentPt: column.wPt,
                });
              })),
            });
          })();
      const constrainedRegions = Object.freeze([
        ...priorRegions.slice(0, -1),
        constrainedPrior,
      ]);
      return Object.freeze({
        ...page,
        accumulator: accumulatePageSectionRegion(
          Object.freeze({ ...page.accumulator, sectionRegions: constrainedRegions }),
          pageRegion(
            nextOwner,
            flow.pageIndex,
            pageInterval,
            flow.regionStartBlockPt,
            event.columnSubset,
          ),
        ),
      });
    },
  };
}

function acceptNode(
  state: BodyPaginationState,
  retained: ParagraphLayout | TableLayout,
  source: SourceRef,
  blockExtentPt: number,
  fragmentStartKey: string,
  acceptedOccurrenceIds: Set<string>,
  allocations: BodyFlowAllocation[],
  placement?: Readonly<{
    coordinateSpace: 'logical-body' | 'upright-physical';
    xPt: number;
    yPt: number;
    sectionFlowOwnership?: 'host-flow' | 'page';
  }>,
): BodyPaginationState {
  const page = state.pages.at(-1);
  if (!page || page.kind !== 'content') throw new Error('Body content requires an active page');
  const region = activeRegion(state);
  const columnIndexes = region.columnIndexes
    ?? region.section.columns.map((_, index) => index);
  const column = region.columns[columnIndexes.indexOf(state.flow.columnIndex)]!;
  const flowDomainId = bodyFlowDomainId(state.flow.pageIndex, region.id, state.flow.columnIndex);
  const occurrenceId = bodyOccurrenceKey(source, flowDomainId, fragmentStartKey);
  if (acceptedOccurrenceIds.has(occurrenceId)) {
    throw new Error(`Duplicate body occurrence acceptance: ${occurrenceId}`);
  }
  acceptedOccurrenceIds.add(occurrenceId);
  const projected = projectBodyOccurrence(retained, {
    occurrenceId,
    destination: {
      coordinateSpace: 'logical-page-points',
      flowDomainId,
      translation: {
        // Ordinary table acquisition owns the complete inline placement:
        // physical jc alignment followed by the acquired leading-axis translation. Move
        // that acquisition-local frame into the page column without
        // normalizing its retained X origin away. Explicit out-of-flow
        // placements and paragraphs continue to own an exact destination.
        xPt: placement
          ? placement.xPt - retained.flowBounds.xPt
          : retained.kind === 'table'
            ? column.inlineStartPt
            : column.inlineStartPt - retained.flowBounds.xPt,
        yPt: (placement?.yPt ?? state.flow.cursorBlockPt) - retained.flowBounds.yPt,
      },
    },
  });
  const ownershipRetained = placement?.sectionFlowOwnership === undefined
    ? projected
    : Object.freeze({ ...projected, sectionFlowOwnership: placement.sectionFlowOwnership });
  const admittedBlockStartPt = placement?.yPt ?? state.flow.cursorBlockPt;
  const contentOwned = ownershipRetained.kind === 'paragraph' && ownershipRetained.ordinaryFlow
    ? (() => {
        const contentStartPt = admittedBlockStartPt + ownershipRetained.spacing.beforePt;
        const contentEndPt = admittedBlockStartPt
          + blockExtentPt
          - ownershipRetained.spacing.afterPt;
        return Object.freeze({
          ...ownershipRetained,
          flowBounds: Object.freeze({
            ...ownershipRetained.flowBounds,
            yPt: contentStartPt,
            // Derive both edges from the admitted allocation before retaining
            // the extent. Spacing collapse reuses the same block-end arithmetic,
            // so adjacent content cannot acquire two floating-point boundaries.
            heightPt: Math.max(0, contentEndPt - contentStartPt),
          }),
        });
      })()
    : ownershipRetained;
  const retainedAtDestination = placement?.coordinateSpace === 'upright-physical'
    ? ({
        ...contentOwned,
        ordinaryFlow: false,
        flowBounds: Object.freeze({
          ...contentOwned.flowBounds,
          heightPt: blockExtentPt,
        }),
      } as typeof contentOwned)
    : contentOwned;
  const transition = placeFlowNode(state.flow, retainedAtDestination, blockExtentPt);
  const place = transition.events[0];
  if (!place || place.type !== 'place') throw new Error('Flow placement did not emit an allocation');
  allocations.push(Object.freeze({
    nodeId: retainedAtDestination.id,
    flowDomainId: retainedAtDestination.flowDomainId,
    blockStartPt: place.blockStartPt,
    blockEndPt: place.blockEndPt,
  }));
  const accumulator = accumulatePagePaintNode(page.accumulator, {
    layer: 'body', node: retainedAtDestination,
    ...(placement?.coordinateSpace === 'upright-physical'
      ? { coordinateSpace: 'upright-physical' as const }
      : {}),
  }, true);
  const pages = [...state.pages];
  pages[pages.length - 1] = Object.freeze({ ...page, accumulator });
  return Object.freeze({
    ...state,
    flow: transition.state,
    pages: Object.freeze(pages),
  });
}

/** The wrap geometry the page registry derives from a frame exclusion. */
function frameExclusionFloatRect(carry: FrameCarryObservation): FloatRect {
  const bounds = carry.exclusionBounds;
  return {
    kind: 'frame',
    mode: carry.exclusionMode,
    imageKey: carry.key,
    imageX: carry.bounds.xPt,
    imageY: carry.bounds.yPt,
    imageW: carry.bounds.widthPt,
    imageH: carry.bounds.heightPt,
    xLeft: bounds.xPt,
    xRight: bounds.xPt + bounds.widthPt,
    yTop: bounds.yPt,
    yBottom: bounds.yPt + bounds.heightPt,
    side: 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
    paraId: 0,
  };
}

/**
 * WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: whether a frame's exclusion meets the
 * line bands of ordinary content (paragraph lines, tables) accepted earlier in
 * its flow domain on the active page. Frames themselves never push frames.
 */
function frameMeetsEarlierFlowContent(
  state: BodyPaginationState,
  flowDomainId: string,
  carry: FrameCarryObservation,
): boolean {
  const page = state.pages.at(-1);
  if (!page || page.kind !== 'content') return false;
  // A line occupies its whole column band (its ink extent is not what wraps).
  const column = acquisitionLocation(state).availableBounds;
  const left = carry.exclusionBounds.xPt;
  const right = left + carry.exclusionBounds.widthPt;
  if (column.xPt >= right || column.xPt + column.widthPt <= left) return false;
  const meets = createFloatLineBandPredicate(frameExclusionFloatRect(carry));
  const crosses = (bounds: LayoutRect) => meets(bounds.yPt, bounds.heightPt);
  for (const entry of page.accumulator.paint) {
    if (entry.layer !== 'body') continue;
    const node = entry.node;
    if (node.kind !== 'paragraph' && node.kind !== 'table') continue;
    if (node.flowDomainId !== flowDomainId) continue;
    if (node.kind === 'paragraph') {
      if (!node.ordinaryFlow) continue;
      const bands = node.lines.length > 0 ? node.lines.map((line) => line.bounds) : [node.flowBounds];
      if (bands.some(crosses)) return true;
      continue;
    }
    if (node.ordinaryFlow === false) continue;
    if (crosses(node.flowBounds)) return true;
  }
  return false;
}

function paragraphFragmentStartKey(cursor: ParagraphFragmentCursor): string {
  return cursor.boundary === null
    ? 'root'
    : `paragraph:${cursor.boundary.segIndex}:${cursor.boundary.charOffset}`;
}

function hasFollowingInkContent(input: BodyLayoutInput, startIndex: number): boolean {
  for (let index = startIndex; index < input.sequence.length; index += 1) {
    const entry = input.sequence[index]!;
    if (entry.kind === 'consume-source') continue;
    if (entry.kind === 'authored-break') {
      if (entry.break !== 'lastRenderedPageBreak') return false;
      continue;
    }
    if (entry.kind === 'begin-section') {
      if (entry.section.startType !== 'continuous') return false;
      continue;
    }
    const block = entry.kind === 'adjacent-table-group' ? entry : entry.block;
    if (block.kind !== 'paragraph') return true;
    if (block.pageBreakBefore) return false;
    if (block.inkless !== true) return true;
  }
  return false;
}

function isUndecoratedInklessMark(layout: ParagraphLayout): boolean {
  return layout.paragraphMark !== undefined
    && layout.lines.length === 0
    && layout.shading === undefined
    && layout.borders.length === 0
    && layout.resources.length === 0
    && layout.drawings.length === 0
    && layout.textBoxes.length === 0;
}

function locationAfter(
  location: BodyAcquisitionLocation,
  blockExtentPt: number,
): BodyAcquisitionLocation {
  const yPt = location.cursorPt.yPt + blockExtentPt;
  const blockEndPt = location.availableBounds.yPt + location.availableBounds.heightPt;
  return Object.freeze({
    ...location,
    cursorPt: Object.freeze({ ...location.cursorPt, yPt }),
    availableBounds: Object.freeze({
      ...location.availableBounds,
      yPt,
      heightPt: Math.max(0, blockEndPt - yPt),
    }),
  });
}

function sectionContentPage(draft: CanonicalPageDraft) {
  return {
    pageIndex: draft.accumulator.pageIndex,
    sectionRegions: draft.accumulator.sectionRegions.map((region) => ({
      sectionOccurrenceId: region.sectionOccurrenceId,
      flowDomainIds: (region.columnIndexes ?? region.section.columns.map((_, index) => index))
        .map((columnIndex) => bodyFlowDomainId(draft.accumulator.pageIndex, region.id, columnIndex)),
    })),
    contentFlowDomainIds: draft.accumulator.readingOrder.map((node) => node.flowDomainId),
  };
}

/** Single authority for page numbering, ownership and story-slot context at
 * opening and final observation/composition. §17.6.12, §17.10.1/.6.
 * Compatibility rule: word-continuous-section-page-number-restart anchors
 * a restart to retained section content,
 * including shared pages, but excludes transition-only empty regions. */
function resolvePageStoryContext(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  previousNumber: LayoutPage['pageNumber'] | undefined,
  firstContentPageIndex: number | undefined,
  parityBlank: boolean,
) {
  const firstPageOfSection = previousNumber?.sectionOccurrenceId !== owner.sectionOccurrenceId;
  const displayNumber = firstPageOfSection && owner.pageNumbering.start !== null
    ? wordContinuousSectionRestartDisplayNumber(owner.pageNumbering.start, pageIndex, firstContentPageIndex ?? pageIndex)
    : (previousNumber?.displayNumber ?? 0) + 1;
  return {
    owner,
    pageIndex,
    section: sectionContextForPage(owner, pageIndex),
    firstPageOfSection,
    parityBlank,
    firstContentPageIndex,
    pageNumber: { displayNumber, format: owner.pageNumbering.format ?? 'decimal',
      sectionOccurrenceId: owner.sectionOccurrenceId },
  };
}

type PageStoryContext = ReturnType<typeof resolvePageStoryContext>;

function finalize(state: BodyPaginationState, owners: ReadonlyMap<string, BodySectionLayoutInput>) {
  const contentFirstAppearance = sectionContentFirstAppearancePageIndices(state.pages.map(sectionContentPage));
  let previousNumber: LayoutPage['pageNumber'] | undefined;
  const pageStoryContexts: PageStoryContext[] = [];
  const pages = state.pages.map((draft) => {
    const owner = owners.get(draft.accumulator.sectionOccurrenceId)!;
    const context = resolvePageStoryContext(owner, draft.accumulator.pageIndex,
      previousNumber, contentFirstAppearance.get(owner.sectionOccurrenceId), draft.kind === 'parity-blank');
    const { pageNumber, firstPageOfSection: firstSectionOwnedPage } = context;
    previousNumber = pageNumber;
    pageStoryContexts.push(context);
    if (draft.kind === 'parity-blank') {
      return createParityBlankLayoutPage({
        pageIndex: draft.accumulator.pageIndex,
        physicalPage: draft.accumulator.physicalPage,
        sectionOccurrenceId: draft.accumulator.sectionOccurrenceId,
        section: draft.accumulator.section,
        pageBorders: draft.accumulator.pageBorders,
        firstSectionOwnedPage,
        pageNumber,
      });
    }
    return finalizeLayoutPage(draft.accumulator, pageNumber, firstSectionOwnedPage);
  });
  const visited = new WeakSet<object>();
  const diagnostics = pages.flatMap((page) =>
    pageLayerNodes(page).flatMap(({ node }) => nestedStoryDiagnostics(node, visited)));
  const layout: DocumentLayout = { pages, diagnostics };
  // Convergence candidates are private to this synchronous pagination call.
  // Validating and deep-freezing the full document graph here repeated that
  // O(document) boundary for every anchor/header/footer pass. The accepted
  // composed layout crosses the invariant/freeze boundary exactly once below.
  return { layout, pageStoryContexts: Object.freeze(pageStoryContexts) };
}

/** One completed body-pagination pass. */
export type BodyPaginationPassResult = Readonly<{
  layout: DocumentLayout;
  session: BodyLayoutSession;
  /** Finalization resolves these with the same authority used at page opening. */
  pageStoryContexts: readonly PageStoryContext[];
  allocations: readonly BodyFlowAllocation[];
  endnoteLayoutsByPage: ReadonlyMap<number, readonly NoteLayout[]>;
  footnoteReserveByPage: ReadonlyMap<number, number>;
  footnoteLayoutsByPage: ReadonlyMap<number, readonly NoteLayout[]>;
  /** Non-zero source-cut shifts of continued tails (FootnotePartition.hostShiftPt). */
  footnoteHostShiftsByPage: ReadonlyMap<number, ReadonlyMap<string, number>>;
  terminalDiagnostic: LayoutDiagnostic | null;
  /** Every read of the anchor-convergence carry, in pass order. */
  anchorInputs: readonly PageAnchorInputEvent[];
  /** `anchorInputs` serialized for exact comparison between passes. */
  serializedAnchorInputs: readonly string[];
  /** Text-anchored frame groups accepted in this pass, in acceptance order
   * (WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH). */
  frameObservations: readonly PlacedFrameObservation[];
  frameAnchorObservations: readonly FrameAnchorObservation[];
}>;

interface BodyPaginationPassObserver {
  shouldPublish(committedPages: number): boolean;
  publish(pass: BodyPaginationPassResult, processedEntries: number): void;
  /** Receives, in pass order, every value the pass reads from the page-anchor
   * convergence carry (see `anchorStablePageLimit`). */
  onPageAnchorInput?(event: PageAnchorInputEvent): void;
  /** Whether the applied anchor plan is still being discovered through this
   * page. The pass suspends (yields) before reading the plan there until it
   * is not, so whoever drives it can make progress elsewhere meanwhile. */
  pageAnchorPlanPending?(pageIndex: number): boolean;
  /** Receives the committed page drafts at each suspension point; all but
   * the last (live) one are closed. */
  onPageDrafts?(pages: readonly CanonicalPageDraft[]): void;
}

type PageStartAnchors = PageAnchorPrescanInput['anchors'];

/** Receives provisional snapshots of one or more passes of the canonical
 * session; `pageIndexLimit` bounds the leading pages the caller may show. */
interface BodyPagePublisher {
  /** Pages already delivered; a publication must extend this prefix. */
  readonly publishedPages: number;
  readonly failed: boolean;
  publish(pass: BodyPaginationPassResult, processedEntries: number, pageIndexLimit: number): void;
}

/** Checkpoint schedule for one pass: publish when the committed page count
 * doubles, and only while the pass can still extend the published prefix. */
function passPublicationObserver(
  publisher: BodyPagePublisher,
  pageIndexLimit: (pass: BodyPaginationPassResult) => number,
  extra: Readonly<{
    onPageAnchorInput?: (event: PageAnchorInputEvent) => void;
    canExtend?: () => boolean;
  }> = {},
): BodyPaginationPassObserver {
  let nextCheckpoint = 1;
  return {
    // One committed page beyond the published prefix is the live page, which
    // is never published; a checkpoint needs at least one more.
    shouldPublish: (committedPages) => (
      !publisher.failed
      && committedPages >= Math.max(nextCheckpoint, publisher.publishedPages + 2)
      && (extra.canExtend?.() ?? true)
    ),
    publish: (pass, processedEntries) => {
      const pages = pass.layout.pages.length;
      nextCheckpoint = Math.max(pages + 1, pages * 2);
      publisher.publish(pass, processedEntries, pageIndexLimit(pass));
    },
    ...(extra.onPageAnchorInput ? { onPageAnchorInput: extra.onPageAnchorInput } : {}),
  };
}

/** The page index of a pass snapshot's live (last, still open) page. */
function livePageIndex(pass: BodyPaginationPassResult): number {
  return pass.layout.pages.at(-1)?.pageIndex ?? 0;
}

/** Optional observer for paintable snapshots of the canonical pagination
 * session: the seed pass, or for a document with page-owned anchors each pass
 * of the unseeded anchor convergence, limited to pages that convergence can no
 * longer change. The generator continues from the same suspended state after
 * every publication; no source prefix is replayed. */
export interface BodyPaginationObserver {
  onPages(layout: DocumentLayout, processedEntries: number): void;
}

function paginationPassResult(
  state: BodyPaginationState,
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
  session: BodyLayoutSession,
  allocations: readonly BodyFlowAllocation[],
  footnoteReserveByPage: ReadonlyMap<number, number>,
  footnoteLayoutsByPage: ReadonlyMap<number, readonly NoteLayout[]>,
  footnoteHostShiftsByPage: ReadonlyMap<number, ReadonlyMap<string, number>>,
  terminalDiagnostic: LayoutDiagnostic | null,
  anchorInputs: readonly PageAnchorInputEvent[] = [],
  serializedAnchorInputs: readonly string[] = [],
  frameObservations: readonly PlacedFrameObservation[] = [],
  frameAnchorObservations: readonly FrameAnchorObservation[] = [],
  endnoteLayoutsByPage: ReadonlyMap<number, readonly NoteLayout[]> = new Map(),
): BodyPaginationPassResult {
  const { layout, pageStoryContexts } = finalize(state, owners);
  const retainedPageIndexes = new Set(layout.pages.map((page) => page.pageIndex));
  const retainedNodeIds = new Set(layout.pages.flatMap((page) => (
    pageLayerNodes(page).map(({ node }) => node.id)
  )));
  // Retained node ids extend their body occurrence key (projectBodyOccurrence).
  const retainedOccurrenceIds = new Set([...retainedNodeIds].map((id) => id.split('/node/')[0]!));
  return Object.freeze({
    layout,
    session,
    pageStoryContexts,
    allocations: Object.freeze(allocations.filter((allocation) => (
      retainedNodeIds.has(allocation.nodeId)
    ))),
    endnoteLayoutsByPage: new Map([...endnoteLayoutsByPage]
      .filter(([pageIndex]) => retainedPageIndexes.has(pageIndex))),
    footnoteReserveByPage: new Map([...footnoteReserveByPage]
      .filter(([pageIndex]) => retainedPageIndexes.has(pageIndex))),
    footnoteLayoutsByPage: new Map([...footnoteLayoutsByPage]
      .filter(([pageIndex]) => retainedPageIndexes.has(pageIndex))),
    footnoteHostShiftsByPage: new Map([...footnoteHostShiftsByPage]
      .filter(([pageIndex]) => retainedPageIndexes.has(pageIndex))
      .map(([pageIndex, shifts]) => [pageIndex, new Map(shifts)])),
    terminalDiagnostic,
    anchorInputs: Object.freeze([...anchorInputs]),
    serializedAnchorInputs: Object.freeze([...serializedAnchorInputs]),
    // An acceptance a later transaction replaced leaves no retained node.
    frameObservations: Object.freeze(frameObservations.filter((observation) => (
      retainedOccurrenceIds.has(observation.nodeId)
    ))),
    frameAnchorObservations: Object.freeze([...frameAnchorObservations]),
  });
}

/**
 * A pagination computation that can be driven one body entry at a time.
 *
 * Every pass in this module is written as a generator so that the SAME code can
 * run to completion synchronously (`drainPagination`) or be spread across event-
 * loop turns (`drainPaginationAsync` in `body-paginator-async.ts`). Suspending
 * between body entries is safe because an entry boundary is the one point where
 * the pass holds no half-applied state: the immutable `BodyPaginationState` has
 * just been committed, and the kernel session's acquisition cursor sits between
 * blocks. A generator — rather than hoisting the pass's locals into an explicit
 * context object — keeps that state exactly where it already lives, so the
 * resumable and non-resumable paths cannot drift apart.
 *
 * The yielded value is the number of pages committed so far, which is what
 * progressive consumers watch.
 */
export type PaginationSteps<T> = Generator<number, T, void>;

/** Run a pagination generator straight through, ignoring the step boundaries. */
export function drainPagination<T>(steps: PaginationSteps<T>): T {
  let step = steps.next();
  while (!step.done) step = steps.next();
  return step.value;
}

type PageWrapDestination = Readonly<{
  kind: 'drawing';
  occurrenceId: string;
  paragraphSource: SourceRef;
  pageIndex: number;
  flowDomainId: string;
}> | Readonly<{
  kind: 'floating-table';
  occurrenceId: string;
  tableSource: SourceRef;
  bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>;
  pageIndex: number;
  flowDomainId: string;
}> | Readonly<{
  kind: 'host-drawing';
  occurrenceId: string;
  paragraphSource: SourceRef;
  pageIndex: number;
  flowDomainId: string;
  carry: CarriedHostAnchor;
}> | Readonly<{
  kind: 'frame';
  occurrenceId: string;
  paragraphSource: SourceRef;
  pageIndex: number;
  flowDomainId: string;
  carry: FrameCarryObservation;
}> | Readonly<{
  /** WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: the frames of `runKey` may not be
   * placed in this flow domain, because their anchor paragraph started in a
   * later one. Never prescanned. */
  kind: 'frame-deferral';
  occurrenceId: string;
  runKey: string;
  pageIndex: number;
  flowDomainId: string;
}>;

/** Where the anchor paragraph of a frame run started, against its frames. */
type FrameAnchorObservation = Readonly<{
  runKey: string;
  framesPageIndex: number;
  framesFlowDomainId: string;
  /** The frames' flow domain held earlier content, so moving them can help. */
  framesAfterContent: boolean;
  anchorPageIndex: number;
  anchorFlowDomainId: string;
}>;

/** A text-anchored frame group as accepted in one pass, with whether its
 * exclusion meets text laid out before it in its flow domain on that page. */
type PlacedFrameObservation = Readonly<{
  nodeId: string;
  pageIndex: number;
  flowDomainId: string;
  meetsEarlierContent: boolean;
  carry: FrameCarryObservation;
}>;

function* paginateBodyPassSteps(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
  anchorDestinations: ReadonlyMap<string, PageWrapDestination> | null,
  minimumTablePageBySource: ReadonlyMap<string, number> | null,
  balancePlan: BodyBalancePlan,
  observer?: BodyPaginationPassObserver,
  anchorLineDeferrals: AnchorLineDeferrals | null = null,
): PaginationSteps<BodyPaginationPassResult> {
  const kernel = bodyLayoutKernelOf(services);
  if (!kernel) throw new Error('Body layout kernel is not attached to the supplied services');
  const owners = ownerMap(input);
  // WORD_PAGE_WIDE_FOOTNOTE_AREA: only the live page replays when its common area changes.
  const pageLocalNotes = input.compatibilityMode === 15
    && (input.initialSection.context.columns.length > 1 || input.sequence.some(entry =>
      entry.kind === 'begin-section' && entry.section.context.columns.length > 1));
  const acceptedOccurrenceIds = new Set<string>();
  const allocations: BodyFlowAllocation[] = [];
  let reserveSession: BodyLayoutSession | undefined;
  let previousPageContext: PageStoryContext | undefined;
  const contentFirstAppearance = new Map<string, number>();
  let nextContentPageIndex = 0;
  const measuredOpeningReserves = new Map<number, HeaderFooterReserve>();
  // ECMA-376 §17.6.11: when a repeat pass opens an unseen physical page,
  // measure its selected story with that page's owner/slot/number. Discovery
  // stays observation-free; no other page's reserve or note plan is borrowed.
  // Only openings use this lookup. Preflight/admission keeps the previous
  // pass's observations (zero for unseen pages), preserving its fit decisions;
  // exact final reserve convergence still validates the resulting pages.
  const reserveForOpening = (owner: BodySectionLayoutInput, pageIndex: number, parityBlank = false, automaticOpening = false): HeaderFooterReserve => {
    // A newly measured page may be smaller than the last complete pass's
    // unseen-page fit input. If overflow leaves that physical page empty,
    // observe the geometry already opened before trying another opening.
    // Retry the reserve pass; keep every fit/admission consumer on
    // the previous-observation contract. Authored blanks and pages that advanced
    // body or incoming notes are not failed openings. No footnote plan is minted.
    // Exact numbering does not remove this case: an indivisible row with a note
    // can fit the unseen-page preflight but not its actual first/default band.
    // Without refinement the cantSplit production regressions exhaust 10,000
    // empty pages. Replay shares the hard ledger; it cannot mint work credits.
    if (automaticOpening && !state.flow.pageHasContent && state.footnoteReservePt === 0
      && measuredOpeningReserves.has(state.flow.pageIndex)) {
      throw new OpenedPageReserveObservation(Object.freeze(Array.from(
        { length: Math.max(reserves.length, state.flow.pageIndex + 1) },
        (_, index) => measuredOpeningReserves.get(index) ?? reserves[index] ?? Object.freeze({ top: 0, bottom: 0 }),
      )));
    }
    // Every closed content page is scanned once. Reuse finalization's exact
    // content-domain projection; a continuous empty region earns no restart.
    if (pageIndex > 0) {
      while (nextContentPageIndex < state.pages.length) {
        const draft = state.pages[nextContentPageIndex++]!;
        for (const [id, first] of sectionContentFirstAppearancePageIndices([sectionContentPage(draft)])) {
          if (!contentFirstAppearance.has(id)) contentFirstAppearance.set(id, first);
        }
      }
    }
    const observed = reserves[pageIndex];
    // Finalization may discover the content anchor after an authored empty page.
    // Keep that exact observation when the owner matches; current retained
    // content takes precedence when this pass has already established it.
    const firstContent = contentFirstAppearance.get(owner.sectionOccurrenceId)
      ?? (observed?.observedPage?.owner === owner.sectionOccurrenceId
        ? observed.observedPage.firstContentPageIndex : undefined);
    const context = resolvePageStoryContext(owner, pageIndex,
      previousPageContext?.pageNumber, firstContent, parityBlank);
    const { firstPageOfSection, pageNumber } = context;
    previousPageContext = context;
    const matching = observed?.observedPage?.owner === owner.sectionOccurrenceId
      && observed.observedPage.firstPageOfSection === firstPageOfSection
      && observed.observedPage.parityBlank === parityBlank
      && observed.observedPage.pageNumber.displayNumber === pageNumber.displayNumber
      && observed.observedPage.pageNumber.format === pageNumber.format;
    if (parityBlank || reserves.length === 0) return { top: 0, bottom: 0 };
    if (matching) return observed!;
    if (!reserveSession) throw new Error('Unobserved page reserve requires an open story session');
    const reserve = Object.freeze({
      ...measurePageHeaderFooterReserve(reserveSession, context),
      observedPage: { owner: owner.sectionOccurrenceId, firstPageOfSection, parityBlank, pageNumber,
        ...(context.firstContentPageIndex === undefined ? {}
          : { firstContentPageIndex: context.firstContentPageIndex }) },
    });
    measuredOpeningReserves.set(pageIndex, reserve);
    return reserve;
  };
  const initialReserve = reserveForOpening(input.initialSection, 0);
  const initialInterval = pageBodyInterval(input.initialSection, 0, initialReserve);
  const initialFlow = createPageFlowState(flowSection(input.initialSection, 0), {
    pageContentStartBlockPt: initialInterval.blockStartPt,
    pageContentEndBlockPt: initialInterval.blockEndPt,
  });
  let state = createBodyPaginationState(
    initialFlow,
    openDraft(input.initialSection, 0, initialInterval),
  );
  const balanceTargetFor = (target: BodyPaginationState): number | null => {
    const planned = balancePlan.get(target.flow.section.sectionOccurrenceId);
    return planned?.pageIndex === target.flow.pageIndex ? planned.targetPt : null;
  };
  state = setBodyBalanceTarget(state, balanceTargetFor(state));
  const factory = transitionFactory(owners, reserveForOpening);
  // Source entry whose paragraph or keep-with-next set was relocated to a new
  // physical page by automatic overflow. Compatibility rules:
  // word-automatic-paragraph-top-spacing and word-automatic-keep-next-start-spacing.
  let automaticPageStartEntryIndex: number | null = null;
  // Compatibility rule: word-standalone-hard-page-break-top-spacing.
  let standaloneHardPageStartEntryIndex: number | null = null;
  const session = kernel.openBodyLayoutSession({
    source: input.source,
    section: input.initialSection.context,
    initialLocation: acquisitionLocation(state),
  }, services, options);
  reserveSession = session;
  const freshPageExtent = (target: BodyPaginationState): number => {
    const owner = owners.get(target.flow.section.sectionOccurrenceId);
    if (!owner) {
      throw new Error(`Unknown body section ${target.flow.section.sectionOccurrenceId}`);
    }
    const nextPageIndex = target.flow.pageIndex + 1;
    const interval = pageBodyInterval(
      owner,
      nextPageIndex,
      reserves[nextPageIndex] ?? { top: 0, bottom: 0 },
    );
    // A same-page §17.18.77 region may begin below the physical body origin.
    // Fresh-page admission is governed by the next page's complete reserved
    // interval, not by that reduced current-page section band.
    return interval.blockEndPt - interval.blockStartPt;
  };
  const pageStartAnchors = (target: BodyPaginationState, startIndex: number): PageStartAnchors => {
    if (anchorDestinations !== null) {
      const location = acquisitionLocation(target);
      return plannedPageStartAnchors(anchorDestinations, location.pageIndex, location.flowDomainId);
    }
    const anchors: Array<Readonly<{
      kind: 'drawing'; occurrenceId: string; paragraphSource: SourceRef;
    }>> = [];
    for (let index = startIndex; index < input.sequence.length; index += 1) {
      const entry = input.sequence[index]!;
      if (entry.kind === 'authored-break' && entry.break !== 'column') break;
      if (entry.kind === 'begin-section' && entry.section.startType !== 'continuous') break;
      if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') continue;
      if (index > startIndex && entry.block.pageBreakBefore) break;
      entry.block.pageOwnedAnchorOccurrenceIds?.forEach((occurrenceId) => anchors.push(
        Object.freeze({ kind: 'drawing', occurrenceId, paragraphSource: entry.block.source }),
      ));
    }
    return Object.freeze(anchors);
  };
  // Every read of the anchor-convergence carry, in pass order. A later pass
  // proves an anchor-line deferral against these reads, and the progressive
  // observer bounds stable pages with them.
  const anchorInputs: PageAnchorInputEvent[] = [];
  const frameObservations: PlacedFrameObservation[] = [];
  const frameAnchorObservations: FrameAnchorObservation[] = [];
  // The frame run whose anchor paragraph has not been accepted yet.
  let pendingFrameRun = null as Omit<FrameAnchorObservation, 'anchorPageIndex' | 'anchorFlowDomainId'> | null;
  const frameRunDeferred = (runKey: string): boolean => {
    if (!anchorDestinations) return false;
    const at = acquisitionLocation(state);
    const deferral = anchorDestinations.get(frameDeferralKey(runKey, at.pageIndex, at.flowDomainId));
    return deferral?.kind === 'frame-deferral';
  };
  const serializedAnchorInputs: string[] = [];
  let observedAnchorInputs = 0;
  const flushAnchorInputs = () => {
    while (observedAnchorInputs < anchorInputs.length) {
      const event = anchorInputs[observedAnchorInputs++]!;
      observer?.onPageAnchorInput?.(event);
    }
  };
  const recordAnchorInput = (event: PageAnchorInputEvent) => {
    anchorInputs.push(event);
    serializedAnchorInputs.push(serializeAnchorInput(event, pageStartAnchorsIdentity));
    // Page-local note trials may replace their live page. External anchor
    // observers receive only its accepted reads, immediately before closure.
    if (!pageLocalNotes) flushAnchorInputs();
  };
  // Page-owned anchors registered on the current physical page.
  let registeredAnchorPage = -1;
  const registeredAnchorKeys = new Set<string>();
  // Anchor-line deferrals applied on the current physical page.
  let deferralPage = -1;
  const appliedDeferralKeys = new Set<string>();
  const prescanPageAnchors = function* (
    target: BodyPaginationState,
    startIndex: number,
  ): Generator<number, void, void> {
    if (anchorDestinations !== null && observer?.pageAnchorPlanPending) {
      const pageIndex = acquisitionLocation(target).pageIndex;
      while (observer.pageAnchorPlanPending(pageIndex)) yield target.pages.length;
    }
    const anchors = pageStartAnchors(target, startIndex);
    const at = acquisitionLocation(target);
    // Recorded even when empty: a later pass that prescans anchors here
    // would change this flow domain.
    recordAnchorInput(Object.freeze({
      kind: 'prescan',
      pageIndex: at.pageIndex,
      flowDomainId: at.flowDomainId,
      anchors,
    }));
    if (registeredAnchorPage !== at.pageIndex) {
      registeredAnchorPage = at.pageIndex;
      registeredAnchorKeys.clear();
    }
    anchors.forEach((anchor) => registeredAnchorKeys.add(anchor.occurrenceId));
    if (anchors.length === 0) return;
    if (!session.prescanPageAnchors) {
      throw new Error('Page-owned anchors require canonical prescan acquisition');
    }
    const location = acquisitionLocation(target);
    const delta = session.prescanPageAnchors({
      anchors,
      location,
      availableInlineExtentPt: location.availableBounds.widthPt,
    });
    if (delta) session.commitFlowRegistryDelta(delta);
  };
  const deferralContext: AnchorLineDeferralContext = Object.freeze({
    reads: serializedAnchorInputs,
  });
  /** Whether a proven anchor-line deferral (see anchor-line-deferral.ts)
   * sends the line anchoring `keys` past the current physical page. */
  const anchorLineDeferred = (keys: readonly string[]): boolean => {
    if (anchorLineDeferrals === null || keys.length === 0) return false;
    const pageIndex = state.flow.pageIndex;
    if (registeredAnchorPage === pageIndex && keys.some((key) => registeredAnchorKeys.has(key))) {
      return false;
    }
    const keysId = anchorKeysId(keys);
    if (deferralPage === pageIndex && appliedDeferralKeys.has(keysId)) return true;
    let proof: AnchorLineDeferralProof | undefined;
    for (const key of keys) {
      proof = anchorLineDeferrals.get(anchorLineDeferralKey(key, pageIndex));
      if (proof) break;
    }
    if (!proof || !anchorLineDeferralApplies(proof, keys, pageIndex, deferralContext)) return false;
    if (deferralPage !== pageIndex) {
      deferralPage = pageIndex;
      appliedDeferralKeys.clear();
    }
    appliedDeferralKeys.add(keysId);
    recordAnchorInput(Object.freeze({
      kind: 'anchor-line-deferral',
      pageIndex,
      keys: Object.freeze([...keys].sort()),
    }));
    return true;
  };
  /** First line of `layout` that a proven deferral keeps off this page. */
  const deferredAnchorLineIndex = (layout: ParagraphLayout): number | undefined => {
    if (anchorLineDeferrals === null) return undefined;
    const keysByLine = pageOwnedAnchorKeysByLine(layout);
    for (let index = 0; index < keysByLine.length; index += 1) {
      if (anchorLineDeferred(keysByLine[index]!)) return index;
    }
    return undefined;
  };
  let capturePageCheckpoint: (entryIndex: number) => void = () => {};
  let beforeNotePageClose = () => {};
  let resetNotePage = () => {};
  let placePendingFootnoteContinuations: () => void = () => {};
  const commitTransition = function* (
    transition: ReturnType<typeof applyAuthoredBreak>,
    nextEntryIndex: number,
    suppressFirstParagraphSpaceBefore = false,
    skipPageAnchorPrescan = false,
  ): Generator<number, void, void> {
    if (transition.state.pageIndex >= MAX_BODY_LAYOUT_PAGES) {
      throw new Error(`Document page budget exceeded (${MAX_BODY_LAYOUT_PAGES} pages)`);
    }
    const previousPageIndex = state.flow.pageIndex;
    if (transition.state.pageIndex !== previousPageIndex) {
      beforeNotePageClose();
      flushAnchorInputs();
    }
    const previousFlowDomainId = acquisitionLocation(state).flowDomainId;
    const opensAutomaticPage = transition.events.some((event) => (
      event.type === 'next-page' && event.reason === 'overflow'
    ));
    const opensSamePageColumnRegion = transition.events.some((event) => (
      event.type === 'begin-section'
      && 'placement' in event
      && event.placement === 'same-page-column'
    ));
    state = commitPageFlowTransition(state, transition, factory);
    state = setBodyBalanceTarget(state, balanceTargetFor(state));
    if (state.flow.pageIndex !== previousPageIndex) {
      automaticPageStartEntryIndex = opensAutomaticPage && suppressFirstParagraphSpaceBefore
        ? nextEntryIndex
        : null;
      const nextLocation = acquisitionLocation(state);
      session.resetPageAcquisition(nextLocation);
      resetNotePage();
      capturePageCheckpoint(nextEntryIndex);
      if (!skipPageAnchorPrescan) yield* prescanPageAnchors(state, nextEntryIndex);
      placePendingFootnoteContinuations();
    } else {
      const nextLocation = acquisitionLocation(state);
      session.moveAcquisitionCursor(nextLocation);
      // WORD_PER_COLUMN_FOOTNOTE_BANDS: a new column owns its own note band;
      // continuations of the previous column's notes open it.
      if (nextLocation.flowDomainId !== previousFlowDomainId
        && (perColumnFootnoteBands() || pageHasColumnFootnoteBands(state.flow.pageIndex))) {
        state = setBodyFootnoteReserve(state, footnoteReserveForActiveBand());
        if (perColumnFootnoteBands()) placePendingFootnoteContinuations();
      }
      // §17.18.77 keeps the physical page but opens a distinct flow domain.
      // The outgoing source scan intentionally stopped at the section mark, so
      // acquire incoming page-owned wrap authority before its first paragraph.
      if (opensSamePageColumnRegion) {
        yield* prescanPageAnchors(state, nextEntryIndex);
      }
    }
  };
  const footnoteIdsByPage = new Map<number, Set<string>>();
  const footnoteReserveByPage = new Map<number, number>();
  const footnoteLayoutsByPage = new Map<number, NoteLayout[]>();
  // Non-zero source-cut shifts of committed continued tails, by page and note
  // (FootnotePartition.hostShiftPt; footnoteBandPlan).
  const footnoteHostShiftsByPage = new Map<number, Map<string, number>>();
  const pendingFootnoteContinuations = new Map<string, FootnoteCursor>();
  let noteController: PageNoteController | undefined;
  let mode15NoteOwner: BodyAcquisitionLocation | undefined;
  // The normative WW8 override keeps only reference-free following text on
  // this physical page. Mode15 is the measured later-reference exception.
  let continuousSectionNotePage: number | null = null;
  let modernFootnoteAdmission: (ids: readonly string[], mandatoryUnit: boolean, allowPartial?: boolean) => boolean | undefined = () => undefined;
  let openModernIncomingNotes: () => boolean = () => false;
  // WORD_PER_COLUMN_FOOTNOTE_BANDS: below compatibility mode 15 each column of
  // a multi-column region owns the note band of the references it paints;
  // otherwise the page owns one band. Band ids double as note containers.
  const footnoteIdsByBand = new Map<string, Set<string>>();
  const footnoteReserveByBand = new Map<string, number>();
  let pageBandKeys = new Set<string>();
  const retainBandReserve = (band: string, reservePt: number) => {
    pageBandKeys.add(band);
    footnoteReserveByBand.set(band, reservePt);
  };
  const pageFootnoteBandId = (pageIndex: number) => `notes:page:${pageIndex}`;
  const perColumnFootnoteBands = (): boolean => wordPerColumnFootnoteBandsApply(
    input.compatibilityMode,
    activeRegion(state).columns.length,
  );
  const activeFootnoteBandId = (): string => {
    const pageBand = pageFootnoteBandId(state.flow.pageIndex);
    return perColumnFootnoteBands()
      ? `${pageBand}:band:${activeRegion(state).id}:${state.flow.columnIndex}`
      : pageBand;
  };
  // Column band reserves by page, so a flow-domain change reads only its page.
  const columnBandReservesByPage = new Map<number, Map<string, number>>();
  const pageHasColumnFootnoteBands = (pageIndex: number): boolean =>
    columnBandReservesByPage.has(pageIndex);
  /** The reserve of the active flow domain: the page band plus its own
   * column band, or for a page-wide domain below column bands, the deepest. */
  const footnoteReserveForActiveBand = (): number => {
    const pageBand = pageFootnoteBandId(state.flow.pageIndex);
    const wide = footnoteReserveByBand.get(pageBand) ?? 0;
    if (perColumnFootnoteBands()) return wide + (footnoteReserveByBand.get(activeFootnoteBandId()) ?? 0);
    let deepest = 0;
    for (const reservePt of columnBandReservesByPage.get(state.flow.pageIndex)?.values() ?? []) {
      deepest = Math.max(deepest, reservePt);
    }
    return wide + deepest;
  };
  const bandHasFootnotes = (band: string): boolean => (footnoteIdsByBand.get(band)?.size ?? 0) > 0;
  // Pages whose note band opens with a committed incoming tail: real
  // retained lines of a note continued from an earlier page, placed before
  // any body (placePendingFootnoteContinuations). Not a newly referenced note,
  // a notice, a separator-only band or a carried line-free paragraph.
  const incomingTailPages = new Set<number>();
  const partitionFootnote = createFootnotePartitioner();
  // Reserved continuation notice (ECMA-376 §17.18.33). The story
  // is acquired through the ordinary story contract only when a fragment
  // actually continues; its definition is shared, its occurrence page-owned.
  const continuationNotice = input.noteLayoutSettings?.reservedSeparatorRoles?.footnote?.continuationNotice;
  const continuationNoticeFor = (
    location: BodyAcquisitionLocation,
    inlineExtentPt: number,
    pendingCarriesNotice = false,
    independentArea = false,
  ): ContinuationNoticeProvider | undefined => {
    if (!continuationNotice?.paragraph || continuationNotice.paragraph.hidden) return undefined;
    // One notice per page: a page that already reserved it (committed, or
    // in an earlier reference-line plan of the fragment being admitted)
    // lends it to the new band-final note at commit instead of charging a
    // second reserve.
    // A column band (WORD_PER_COLUMN_FOOTNOTE_BANDS) owns its own notice.
    const band = activeFootnoteBandId();
    if (!independentArea && footnoteLayoutsByPage.get(state.flow.pageIndex)
      ?.some((note) => note.trailing && note.flowDomainId === band)) return undefined;
    if (pendingCarriesNotice) return undefined;
    return () => {
      if (!session.layoutStory) throw new Error('A footnote continuation notice requires story layout');
      const bounds = {
        xPt: location.availableBounds.xPt,
        yPt: 0,
        widthPt: inlineExtentPt,
        heightPt: location.section.geometry.pageHeight,
      };
      const story = session.layoutStory({
        source: continuationNotice.root,
        pageIndex: state.flow.pageIndex,
        section: location.section,
        container: { id: `notes:page:${state.flow.pageIndex}:notice`, kind: 'footnote', bounds },
      });
      return noteSeparatorOccurrence(continuationNotice, story, bounds);
    };
  };
  type FootnoteAdmission = Readonly<{
    ids: readonly string[];
    layouts: readonly NoteLayout[];
    reservePt: number;
    continuations: readonly Readonly<{ id: string; cursor: FootnoteCursor }>[];
  }>;
  // §17.18.77 observes committed references even when their measured reserve is zero;
  // footnoteReservePt remains only the page-local geometry charge.
  const hasFootnoteReferenceOnPage = (pageIndex: number): boolean => (
    (footnoteIdsByPage.get(pageIndex)?.size ?? 0) > 0
  );
  const modernAdmissionResult = (ids: readonly string[], accepted: boolean): FootnoteAdmission => Object.freeze({
    ids: Object.freeze([...ids]), layouts: Object.freeze([]), continuations: Object.freeze([]),
    reservePt: accepted ? 0 : Number.MAX_VALUE,
  });
  /** Whole-note admission for every reference in a slice (no partition). */
  const footnoteAdmission = (
    candidate: ParagraphLayout | TableLayout,
    inlineExtentPt: number,
    retainedReferenceIds?: readonly string[],
  ): FootnoteAdmission => {
    const whole = footnoteAdmissionForIds(retainedReferenceIds ?? footnoteIdsInRetainedSlice(candidate), inlineExtentPt);
    // Frames/tables/indivisible marks use their existing whole-note admission.
    const modern = modernFootnoteAdmission(whole.ids, false, false);
    return modern === undefined ? whole : modernAdmissionResult(whole.ids, modern);
  };
  /** `pending` summarises earlier reference-line plans of the fragment being
   * admitted (maintained incrementally by their owner): their notes are not
   * re-acquired, and they already own the page's separator band and any
   * continuation notice. */
  const footnoteAdmissionForIds = (
    retainedReferenceIds: readonly string[],
    inlineExtentPt: number,
    availableReservePt?: number,
    pending: PendingFootnotes = NO_PENDING_FOOTNOTES,
    extent: FootnotePartitionExtent = 'largest',
  ): FootnoteAdmission => {
    const committed = footnoteIdsByPage.get(state.flow.pageIndex) ?? new Set<string>();
    const pendingIds = pending.ids;
    const ids = [...new Set(retainedReferenceIds)]
      .filter((id) => !committed.has(id) && !pendingIds.has(id));
    const location = acquisitionLocation(state);
    if (ids.length > 0 && continuousSectionNotePage === state.flow.pageIndex) {
      return Object.freeze({ ids: Object.freeze(ids), layouts: Object.freeze([]),
        continuations: Object.freeze([]), reservePt: Number.MAX_VALUE });
    }
    if (ids.length > 0) {
      const modern = modernFootnoteAdmission(ids, pending === LATER_UNIT_REFERENCE,
        availableReservePt === undefined ? undefined : true);
      if (modern !== undefined) return modernAdmissionResult(ids, modern);
    }
    if (ids.length > 0 && !session.layoutNotes) {
      throw new Error('Footnote layout requires a note-capable layout session');
    }
    const plannedTopsPt = reserves[state.flow.pageIndex]?.footnoteTopsPt;
    const band = activeFootnoteBandId();
    const noteLocation = mode15NoteOwner?.section.columns.length === 1 ? mode15NoteOwner : location;
    const noteInlineExtentPt = mode15NoteOwner?.section.columns.length === 1
      ? mode15NoteOwner.availableBounds.widthPt : inlineExtentPt;
    const layouts = ids.length === 0 ? Object.freeze([]) : session.layoutNotes!({
      kind: 'footnote',
      referenceIds: Object.freeze(ids),
      pageIndex: state.flow.pageIndex,
      section: noteLocation.section,
      ...(plannedTopsPt ? { plannedTopsPt } : {}),
      container: {
        id: band,
        kind: 'footnote',
        bounds: {
          xPt: noteLocation.availableBounds.xPt,
          yPt: 0,
          widthPt: noteInlineExtentPt,
          heightPt: noteLocation.section.geometry.pageHeight,
        },
      },
      firstOnPage: !bandHasFootnotes(band) && pendingIds.size === 0,
    });
    const continuations: Array<Readonly<{ id: string; cursor: FootnoteCursor }>> = [];
    const admittedLayouts = layouts.map((note, index) => {
      if (services.allowFootnoteContinuation !== true
        || availableReservePt === undefined || index !== layouts.length - 1) return note;
      const precedingPt = layouts.slice(0, index).reduce((sum, preceding) => sum + preceding.advancePt, 0);
      const partition = partitionFootnote(note, null, Math.max(0, availableReservePt - precedingPt),
        extent === 'minimum' ? undefined : continuationNoticeFor(location, inlineExtentPt, pending.carriesNotice),
        extent);
      if (partition?.nextCursor) {
        continuations.push(Object.freeze({ id: note.source.storyInstance, cursor: partition.nextCursor }));
        return partition.fragment;
      }
      return note;
    });
    return Object.freeze({
      ids: Object.freeze(ids),
      layouts: Object.freeze(admittedLayouts),
      reservePt: admittedLayouts.reduce((sum, note) => sum + note.advancePt, 0),
      continuations: Object.freeze(continuations),
    });
  };
  const commitFootnotes = (
    ids: readonly string[],
    layouts: readonly NoteLayout[],
    continuations: FootnoteAdmission['continuations'] = [],
  ) => {
    if (noteController) {
      let actual = footnoteIdsByPage.get(state.flow.pageIndex);
      if (!actual) footnoteIdsByPage.set(state.flow.pageIndex, actual = new Set());
      ids.forEach(id => actual!.add(id));
      const band = pageFootnoteBandId(state.flow.pageIndex);
      let bandIds = footnoteIdsByBand.get(band);
      if (!bandIds) footnoteIdsByBand.set(band, bandIds = new Set());
      ids.forEach(id => bandIds!.add(id));
      pageBandKeys.add(band);
      noteController.committed(ids);
      return;
    }
    let retained = footnoteIdsByPage.get(state.flow.pageIndex);
    if (!retained) {
      retained = new Set<string>();
      footnoteIdsByPage.set(state.flow.pageIndex, retained);
    }
    ids.forEach((id) => retained!.add(id));
    const band = activeFootnoteBandId();
    let bandIds = footnoteIdsByBand.get(band);
    if (!bandIds) footnoteIdsByBand.set(band, bandIds = new Set<string>());
    ids.forEach((id) => bandIds!.add(id));
    const retainedLayouts = footnoteLayoutsByPage.get(state.flow.pageIndex) ?? [];
    // Reusable full-note acquisition may have first been requested during a
    // rejected admission on another page. Intact notes need destination
    // ownership too; fragmented notes already project only their admitted lines.
    const ownedLayouts = services.allowFootnoteContinuation !== true
      ? layouts : layouts.map(destinationOwnedNote);
    retainedLayouts.push(...ownedLayouts);
    // §17.18.33: the page's one continuation notice ends its note band. A note
    // committed after the carrier takes the already-reserved notice; a later
    // continuing partition on this page reserves none (continuationNoticeFor).
    // The notice and its reserve stay inside the band that reserved them.
    const bandIndexes = retainedLayouts.flatMap((note, index) => note.flowDomainId === band ? [index] : []);
    const carriers = bandIndexes.filter((index) => retainedLayouts[index]!.trailing);
    if (carriers.length > 1) throw new Error('A page note band owns exactly one continuation notice');
    const carrierIndex = carriers[0];
    const lastIndex = bandIndexes.at(-1)!;
    if (carrierIndex !== undefined && carrierIndex !== lastIndex) {
      const moved = moveContinuationNotice(retainedLayouts[carrierIndex]!, retainedLayouts[lastIndex]!);
      retainedLayouts[carrierIndex] = moved.carrier;
      retainedLayouts[lastIndex] = moved.receiver;
    }
    footnoteLayoutsByPage.set(state.flow.pageIndex, retainedLayouts);
    // The retained note band sums individual advances in document order.
    // Adding independently summed reference groups changes floating-point
    // association and can disagree with that exact geometry authority.
    state = addPageFootnoteReserves(state, layouts.map(note => note.advancePt));
    let bandReservePt = footnoteReserveByBand.get(band) ?? 0;
    let pageReservePt = footnoteReserveByPage.get(state.flow.pageIndex) ?? 0;
    for (const note of layouts) {
      bandReservePt += note.advancePt;
      pageReservePt += note.advancePt;
    }
    retainBandReserve(band, bandReservePt);
    if (band !== pageFootnoteBandId(state.flow.pageIndex)) {
      let pageBands = columnBandReservesByPage.get(state.flow.pageIndex);
      if (!pageBands) columnBandReservesByPage.set(state.flow.pageIndex, pageBands = new Map());
      pageBands.set(band, bandReservePt);
    }
    // The page's retained note advances in commit order (the page invariant).
    footnoteReserveByPage.set(state.flow.pageIndex, pageReservePt);
    continuations.forEach(({ id, cursor }) => pendingFootnoteContinuations.set(id, cursor));
  };
  /** The deepest committed content a growing note band must not cover: the
   * active column alone for a column band, every column for a page band. */
  const committedFootnoteContentEndPt = (): number => (
    perColumnFootnoteBands() ? state.flow.cursorBlockPt : state.flow.deepestColumnBlockPt
  );
  placePendingFootnoteContinuations = () => {
    if (pendingFootnoteContinuations.size === 0) return;
    if (openModernIncomingNotes()) return;
    const pending = [...pendingFootnoteContinuations];
    pendingFootnoteContinuations.clear();
    const location = acquisitionLocation(state);
    // A band-dependent tail is acquired with this page's planned band, as a
    // head is (footnoteAdmissionForIds); the plan already includes its cut.
    const plannedTopsPt = reserves[state.flow.pageIndex]?.footnoteTopsPt;
    for (const [pendingIndex, [id, cursor]] of pending.entries()) {
      if (!session.layoutNotes) throw new Error('Footnote continuation requires note layout');
      const capacityPt = Math.max(0, activeBlockEndPt(state)
        - state.footnoteReservePt - committedFootnoteContentEndPt());
      const partition = acquireNoteColumnPartition(session, {
        kind: 'footnote',
        referenceIds: Object.freeze([id]),
        pageIndex: state.flow.pageIndex,
        section: location.section,
        ...(plannedTopsPt ? { plannedTopsPt } : {}),
        container: {
          id: activeFootnoteBandId(),
          kind: 'footnote',
          bounds: {
            xPt: location.availableBounds.xPt,
            yPt: 0,
            widthPt: location.availableBounds.widthPt,
            heightPt: location.section.geometry.pageHeight,
          },
        },
        firstOnPage: perColumnFootnoteBands()
          ? !bandHasFootnotes(activeFootnoteBandId())
          : !hasFootnoteReferenceOnPage(state.flow.pageIndex),
        continuing: true,
      }, cursor, capacityPt, partitionFootnote,
        continuationNoticeFor(location, location.availableBounds.widthPt));
      if (!partition) {
        if (!hasFootnoteReferenceOnPage(state.flow.pageIndex)) {
          throw new NoteCapacityExceededError('footnote', state.flow.pageIndex, `notes:page:${state.flow.pageIndex}`);
        }
        // Library continuation policy: an earlier retained fragment can use
        // this page's capacity without making a later source cursor invalid.
        // Keep that cursor and its following peers in reference order for the
        // next page; emit no empty fragment or extra separator/notice reserve.
        // A first fragment that cannot fit the fresh band still fails above,
        // so deferral cannot create an unbounded sequence of empty pages.
        for (let index = pendingIndex; index < pending.length; index += 1) {
          const [deferredId, deferredCursor] = pending[index]!;
          pendingFootnoteContinuations.set(deferredId, deferredCursor);
        }
        break;
      }
      commitFootnotes([id], [partition.fragment], partition.nextCursor
        ? [Object.freeze({ id, cursor: partition.nextCursor })]
        : []);
      if (partition.hostShiftPt !== 0) {
        let shifts = footnoteHostShiftsByPage.get(state.flow.pageIndex);
        if (!shifts) footnoteHostShiftsByPage.set(state.flow.pageIndex, shifts = new Map());
        shifts.set(noteBandKey(partition.fragment.flowDomainId, id), partition.hostShiftPt);
      }
      // The partitioner also carries a line-free paragraph on its own (an
      // empty paragraph keeps its spacing, ownership and cursor), so a
      // non-null fragment can hold no line. The carried-tail class was
      // observed only with real retained note lines; an empty-only, notice or
      // separator-only band does not ground it.
      if (partition.fragment.story.blocks.some(block => block.kind === 'paragraph' && block.lines.length > 0)) {
        incomingTailPages.add(state.flow.pageIndex);
      }
    }
  };
  // §17.11.21 / §17.18.34 assign each note to the physical page that paints
  // its reference. Growing that page-wide band must not clip a deeper column
  // that the immutable paginator has already committed.
  const additionalFootnoteReserveCapacityPt = (): number => Math.max(
    0,
    activeBlockEndPt(state)
      - state.footnoteReservePt
      - committedFootnoteContentEndPt(),
  );
  const footnoteReserveInvadesCommittedPageContent = (reservePt: number): boolean => (
    reservePt > additionalFootnoteReserveCapacityPt()
  );
  // A region is "fresh" only when moving on cannot gain space. A page whose
  // note band already holds committed lines (carried tails of earlier notes,
  // placed before any body) is not: those lines belong to their own notes and
  // shrink page by page, so relocating gains their space and the next page
  // can still progress. Treating such a page as fresh would force a block in
  // with a note it cannot hold, pushing the page-bottom band (§17.11.21) off
  // the page. Continuation-off pages carry no tails, so there this is the
  // existing body-content rule.
  const pageCanGainSpaceByRelocating = (): boolean => (
    state.flow.pageHasContent || state.footnoteReservePt > 0
  );
  const noteContinuationEligibilityByEntry = new Map<number, boolean>();
  const noteContinuationBeforeHardBreak = (entryIndex: number): boolean => {
    const cached = noteContinuationEligibilityByEntry.get(entryIndex);
    if (cached !== undefined) return cached;
    // word-paragraph-footnote-continuation: cache traversed entries so the
    // bounded compatibility lookahead visits each following block once.
    const visited = [entryIndex];
    const finish = (eligible: boolean): boolean => {
      visited.forEach((index) => noteContinuationEligibilityByEntry.set(index, eligible));
      return eligible;
    };
    for (let index = entryIndex + 1; index < input.sequence.length; index += 1) {
      const prior = noteContinuationEligibilityByEntry.get(index);
      if (prior !== undefined) return finish(prior);
      const next = input.sequence[index]!;
      if (wordFootnoteLookaheadBoundary(next)) return finish(true);
      const block = next.kind === 'body-block' ? next.block
        : next.kind === 'adjacent-table-group' ? next : null;

      if (block) {
        const location = acquisitionLocation(state);
        const measurement = session.measureFollowingBlock({
          input: block,
          location,
          availableInlineExtentPt: location.availableBounds.widthPt,
        });
        if ((measurement.fullFootnoteReferenceIds?.length ?? 0) > 0) return finish(false);
      }
      visited.push(index);
    }
    return finish(false);
  };
  /**
   * Reference-line footnote plans for one paragraph placement attempt
   * (word-reference-line-footnote-allocation; continuation mode only).
   *
   * Events and immutability: each line that first references notes not yet
   * on the page is one admission event. Its plan is acquired once and is
   * immutable: a later event can only add reserve. The reserve charged for a
   * candidate ending at line `e` is the sum of the plans of events at or
   * before `e`, so it never decreases as the candidate grows. If an event's
   * plan cannot fit (no real line of its last note, or its preceding new
   * notes cannot stay whole), the plan is charged whole, the candidate
   * through that line fails, and that line is rejected together with every
   * later line; earlier plans, their source cursors and their charge are
   * untouched and are exactly what `admissionThrough` commits.
   *
   * Body admission unit: a plan is sized against the smallest body fragment
   * the paragraph's own line constraints let end at or after its reference
   * line, not against the reference line alone — the whole paragraph for
   * keepLines at its start (ECMA-376 §17.3.1.14), else the first end that
   * leaves neither an orphan nor a widow (§17.3.1.44, the selector's rule).
   * A note head therefore never consumes body space those constraints
   * require. Where the page already has content, that strict unit governs
   * and the selector relocates the paragraph if it does not fit. On a fresh
   * region relocation cannot help, so a unit that cannot fit (the
   * constraint is impossible there) yields to the next weaker unit, ending
   * with the reference line itself; content still progresses.
   *
   * Work frontier: prefix charges come from one exact prefix-advance index
   * (no slice per event), pending ids/notice ownership and cumulative
   * reserves are maintained incrementally, and text-after facts are
   * precomputed. Where the reserve is a monotone prefix sum (horizontal,
   * splittable, no keepLines unit), planning stops at the first event whose
   * own candidate is proven inadmissible (body through its line plus the
   * cumulative reserve exceeds the selector's limit, or the band would
   * invade a committed column) and at an anchor-deferral line limit. Every
   * longer candidate then fails too: its body and reserve can only be
   * larger, and a completed fragment's admission extent is at least that
   * prefix's advance because horizontal admission only discounts authored
   * trailing space. The truncated sum returned for such candidates is a
   * sound lower bound; no plan beyond the frontier is acquired or charged
   * to the footnote work budget.
   */
  type ReferenceLineFootnotePlans = Readonly<{
    reservePt(lineEnd: number): number;
    admissionThrough(lineEnd: number): FootnoteAdmission;
  }>;
  const referenceLineFootnotePlans = (
    layout: ParagraphLayout,
    options: Readonly<{
      indivisible: boolean;
      continuesFromPrevious: boolean;
      entryIndex: number;
      /** Whether moving the paragraph on can gain space (the selector's own
       * relocation flag); false only for a genuinely fresh region. */
      canRelocate: boolean;
      /** keepLines that still governs this fragment (paragraph start). */
      keepLines: boolean;
      widowControl: boolean;
      lineEndLimit: number | undefined;
      /** The selector's block limit, including any trailing-mark allowance. */
      admissionLimitPt: number;
      /** Authored §17.3.1.33 space after, as given to the selector. */
      authoredSpaceAfterPt: number;
      writingMode: WritingMode;
    }>,
  ): ReferenceLineFootnotePlans => {
    const location = acquisitionLocation(state);
    const inlineExtentPt = location.availableBounds.widthPt;
    const availablePt = location.availableBounds.heightPt;
    const freshExtentPt = freshPageExtent(state);
    const canRelocate = options.canRelocate;
    // Carried-tail class: this page's band already opens with real retained
    // lines of an earlier note (committed before any body). A page fact,
    // independent of capacity, so the real plans and the mandatory-unit minima
    // read the same value. Heads are still sized from `availablePt`, which
    // already excludes that committed tail reserve exactly once.
    const pageOpensWithIncomingTail = incomingTailPages.has(state.flow.pageIndex);
    const total = layout.lines.length;
    const frontierBounded = !options.indivisible && !options.keepLines
      && options.writingMode !== 'vertical-rl';
    const prefixAdvancesPt = options.indivisible
      ? null
      : paragraphContinuingPrefixAdvancesPt(layout, options.continuesFromPrevious);
    // A unit that completes the paragraph is charged with its retained advance,
    // including authored spacing after: §17.3.1.33 spacing follows the last
    // line, and here the note band (§17.11) follows the body, so the body ends
    // at the band rather than at a page edge. The selector's page/column-edge
    // trailing-space admission (WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION) has no
    // evidence at that boundary, so plans do not use it. Public Word 16.113.3
    // controls with exact lines split the note at the same point.
    const bodyAdvancePt = (lineEnd: number): number => prefixAdvancesPt === null
      ? layout.advancePt
      : lineEnd < total ? prefixAdvancesPt[lineEnd]! : prefixAdvancesPt[total]! + layout.spacing.afterPt;
    // The completing body without only its admissible authored after, through
    // the selector's own trailing-space projection. It never sizes a head; it
    // only identifies the after-only class for wordParagraphFootnoteMayContinue.
    const completingBodyWithoutAfterPt = wordFinalParagraphAdmissionExtentPt({
      advancePt: bodyAdvancePt(total),
      retainedSpaceAfterPt: layout.spacing.afterPt,
      authoredSpaceAfterPt: options.authoredSpaceAfterPt,
    });
    // Text after an event's first reference, without suffix copies: a
    // custom-mark reference is a zero-width anchor host and is still the
    // event's reference (§17.18.34 ownership), not following text.
    const isFollowingText = (placement: ParagraphLayout['lines'][number]['placements'][number]) =>
      placement.kind === 'text' && placement.noteReference?.kind !== 'footnote'
      && placement.text.trim().length > 0;
    const textFromLine = new Array<boolean>(total + 1).fill(false);
    for (let line = total - 1; line >= 0; line -= 1) {
      textFromLine[line] = textFromLine[line + 1]! || layout.lines[line]!.placements.some(isFollowingText);
    }
    const textAfterReference = (line: number, ids: ReadonlySet<string>): boolean => {
      const placements = layout.lines[line]?.placements ?? [];
      const first = placements.findIndex((placement) =>
        (placement.kind === 'text' || placement.kind === 'anchor-host')
        && placement.noteReference?.kind === 'footnote' && ids.has(placement.noteReference.id));
      if (first < 0) return false;
      return placements.slice(first + 1).some(isFollowingText) || textFromLine[line + 1]!;
    };
    type FootnoteEvent = Readonly<{ referenceLine: number; lineEnd: number; ids: readonly string[] }>;
    const events: FootnoteEvent[] = [];
    const seen = new Set<string>();
    layout.lines.forEach((line, index) => {
      const ids = footnoteIdsInRetainedLines([line]).filter((id) => !seen.has(id));
      ids.forEach((id) => seen.add(id));
      if (ids.length === 0) return;
      if (options.indivisible && events.length > 0) {
        // An indivisible fragment is admitted whole: one event at its end.
        const merged = events.pop()!;
        events.push({ ...merged, ids: [...merged.ids, ...ids] });
      } else {
        events.push({ referenceLine: index, lineEnd: options.indivisible ? total : index + 1, ids });
      }
    });
    const eventsThrough = new Array<number>(total + 1).fill(0);
    for (const event of events) eventsThrough[event.lineEnd] = (eventsThrough[event.lineEnd] ?? 0) + 1;
    for (let lineEnd = 1; lineEnd <= total; lineEnd += 1) eventsThrough[lineEnd]! += eventsThrough[lineEnd - 1]!;
    const admissionUnitEnds = (lineEnd: number): readonly number[] => {
      if (options.indivisible) return [total];
      const ends: number[] = [];
      if (options.keepLines) ends.push(total);
      if (options.widowControl) ends.push(widowLegalLineEnd(lineEnd, total));
      ends.push(lineEnd);
      return [...new Set(ends)];
    };
    const plans: FootnoteAdmission[] = [];
    const cumulativePt = [0];
    const pending = { ids: new Set<string>(), carriesNotice: false };
    let frontierReached = false;
    // Mandatory-unit lookahead. When an event's body unit also contains later
    // reference lines, those references must be admitted with it, so its
    // plan first reserves their legal minimum: every preceding new note of
    // that event whole and its last new note's first real line (the
    // partitioner's minimum-head extent) when existing continuation policy
    // lets it split, else that note whole. These minima are taken from the
    // measured stories once per event and kept as prefix sums; they are
    // conservative (continuation eligibility is judged with the smallest
    // possible charge, so a note that might have to stay whole is reserved
    // whole). The page's single reserved notice is added once if any of those
    // minima continues and neither the page nor an accepted plan owns it;
    // the current plan then partitions inside the capacity left after that
    // reserve and, if it continues, carries the same notice, so the notice is
    // never reserved twice while planning or committing (one band, one
    // notice). Only the events inside the unit are examined, so ordinary units have
    // no lookahead. The later event is then planned for real against the
    // capacity left, greedily, and earlier accepted plans never change.
    const minimumPt: number[] = [];
    const minimumPrefixPt = [0];
    const minimumContinuesPrefix = [0];
    const requireMinimum = (count: number): void => {
      while (minimumPt.length < count) {
        const index = minimumPt.length;
        const event = events[index]!;
        const whole = footnoteAdmissionForIds(event.ids, inlineExtentPt, undefined, LATER_UNIT_REFERENCE);
        const textAfter = textAfterReference(event.referenceLine, new Set(event.ids));
        const wholeNoteFitsFreshPage = bodyAdvancePt(admissionUnitEnds(event.lineEnd)[0]!)
          + whole.reservePt <= freshExtentPt;
        const mayPartition = wordParagraphFootnoteMayContinue({
          wholeNoteFitsWithReference: false,
          wholeNoteFitsFreshPage,
          textAfterReference: textAfter,
          hardBreakBeforeNextNote: wholeNoteFitsFreshPage && !textAfter && index + 1 === events.length
            && noteContinuationBeforeHardBreak(options.entryIndex),
          // The after-only permission depends on the capacity this later
          // event will actually see, which a minimum taken before its earlier
          // plans cannot know. Where it would be the only permission, the
          // legal minimum stays the whole note (conservative for that branch
          // only); the real plan may still continue the note within whatever
          // that minimum left.
          wholeNoteFitsOnlyWithoutTrailingAfter: false,
          columnBand: perColumnFootnoteBands(),
          // Unlike the after-only class this permission needs no capacity,
          // so the legal minimum uses it like the text/break permissions:
          // the later note's first real line, with any notice in the union.
          pageOpensWithIncomingTail,
        });
        let minimum = mayPartition
          ? footnoteAdmissionForIds(event.ids, inlineExtentPt, 0, LATER_UNIT_REFERENCE, 'minimum')
          : whole;
        const modern = modernFootnoteAdmission(event.ids, true, mayPartition);
        if (modern !== undefined) minimum = modernAdmissionResult(event.ids, modern);
        minimumPt.push(minimum.reservePt);
        minimumPrefixPt.push(minimumPrefixPt[index]! + minimum.reservePt);
        minimumContinuesPrefix.push(minimumContinuesPrefix[index]! + (minimum.continuations.length > 0 ? 1 : 0));
      }
    };
    let noticeAdvancePt: number | undefined;
    const NO_LATER_RESERVE = Object.freeze({ minimaPt: 0, noticePt: 0 });
    const laterUnitReserve = (eventIndex: number, unitEnd: number): Readonly<{ minimaPt: number; noticePt: number }> => {
      const end = eventsThrough[unitEnd]!;
      if (end <= eventIndex + 1) return NO_LATER_RESERVE;
      requireMinimum(end);
      const continues = minimumContinuesPrefix[end]! - minimumContinuesPrefix[eventIndex + 1]! > 0;
      const notice = continues ? continuationNoticeFor(location, inlineExtentPt, pending.carriesNotice) : undefined;
      if (notice) noticeAdvancePt ??= notice()?.advancePt ?? 0;
      return {
        minimaPt: minimumPrefixPt[end]! - minimumPrefixPt[eventIndex + 1]!,
        noticePt: notice ? noticeAdvancePt ?? 0 : 0,
      };
    };
    const planNext = (): void => {
      const event = events[plans.length]!;
      if (frontierBounded && options.lineEndLimit !== undefined && event.lineEnd > options.lineEndLimit) {
        frontierReached = true;
        return;
      }
      const plannedPt = cumulativePt[plans.length]!;
      const whole = footnoteAdmissionForIds(event.ids, inlineExtentPt, undefined, pending);
      const textAfter = textAfterReference(event.referenceLine, new Set(event.ids));
      const laterEvent = plans.length + 1 < events.length;
      let chosen: FootnoteAdmission | undefined;
      for (const unitEnd of admissionUnitEnds(event.lineEnd)) {
        const bodyPt = bodyAdvancePt(unitEnd);
        // Mandatory later references of this unit are charged before this
        // plan takes any remaining capacity. When their reserve already holds
        // the page's single notice, this plan partitions inside the capacity
        // left after it and, if it continues, carries that same notice
        // instead of reserving another one.
        const later = laterUnitReserve(plans.length, unitEnd);
        const laterPt = later.minimaPt + later.noticePt;
        const wholeChargePt = bodyPt + plannedPt + whole.reservePt + laterPt;
        const wholeNoteFitsWithReference = wholeChargePt <= availablePt
          && !footnoteReserveInvadesCommittedPageContent(plannedPt + whole.reservePt + laterPt);
        const wholeNoteFitsFreshPage = wholeChargePt <= freshExtentPt;
        // After-only class: a unit that completes the paragraph misses only
        // because of its admissible authored after. The selector would admit
        // that whole note with the after overlapping the band; permitting a
        // split lets a head of at least one real line be sized beside the full
        // charged body below, keeping that after clear of the band.
        // Known limitation (owner-pending, not Word's rule): when no real line
        // fits there (a one-line note, or after filling the capacity), the
        // partition is null and the existing whole-note fallback still admits
        // the note over that after. In the public one-line control Word instead
        // moved the final reference-bearing line and its note to the next page.
        const wholeNoteFitsOnlyWithoutTrailingAfter = !wholeNoteFitsWithReference
          && unitEnd === total
          && unitEnd <= (options.lineEndLimit ?? total)
          && completingBodyWithoutAfterPt < bodyPt
          && completingBodyWithoutAfterPt + plannedPt + whole.reservePt + laterPt <= availablePt
          && !footnoteReserveInvadesCommittedPageContent(plannedPt + whole.reservePt + laterPt);
        const mayPartition = wordParagraphFootnoteMayContinue({
          wholeNoteFitsWithReference,
          wholeNoteFitsFreshPage,
          textAfterReference: textAfter,
          hardBreakBeforeNextNote: !wholeNoteFitsWithReference && wholeNoteFitsFreshPage
            && !textAfter && !laterEvent && noteContinuationBeforeHardBreak(options.entryIndex),
          wholeNoteFitsOnlyWithoutTrailingAfter,
          pageOpensWithIncomingTail,
          columnBand: perColumnFootnoteBands(),
        });
        chosen = mayPartition
          ? footnoteAdmissionForIds(event.ids, inlineExtentPt, Math.max(0, Math.min(
              availablePt - bodyPt,
              additionalFootnoteReserveCapacityPt(),
            ) - plannedPt - laterPt), pending,
            later.noticePt > 0 ? 'largest-with-reserved-notice' : 'largest')
          : whole;
        const modern = modernFootnoteAdmission(event.ids, false, mayPartition);
        if (modern !== undefined) chosen = modernAdmissionResult(event.ids, modern);
        // The notice is charged once: either in the later reserve or, when
        // this plan carries it, in this plan's own advance.
        const sharedNoticePt = later.noticePt > 0 && chosen.layouts.some((note) => note.trailing)
          ? later.noticePt : 0;
        const unitReservePt = plannedPt + chosen.reservePt + laterPt - sharedNoticePt;
        const unitFits = unitEnd <= (options.lineEndLimit ?? total)
          && bodyPt + unitReservePt <= availablePt
          && !footnoteReserveInvadesCommittedPageContent(unitReservePt);
        if (canRelocate || unitFits) break;
      }
      plans.push(chosen!);
      cumulativePt.push(plannedPt + chosen!.reservePt);
      for (const id of chosen!.ids) pending.ids.add(id);
      if (chosen!.layouts.some((note) => note.trailing)) pending.carriesNotice = true;
      const cumulative = cumulativePt[plans.length]!;
      if (frontierBounded && event.lineEnd < total && (
        bodyAdvancePt(event.lineEnd) + cumulative > options.admissionLimitPt
        || footnoteReserveInvadesCommittedPageContent(cumulative)
      )) frontierReached = true;
    };
    const planEvents = (count: number): void => {
      while (plans.length < count && !frontierReached) planNext();
    };
    return Object.freeze({
      reservePt: (lineEnd: number) => {
        const count = eventsThrough[Math.min(lineEnd, total)]!;
        planEvents(count);
        return cumulativePt[Math.min(count, plans.length)]!;
      },
      admissionThrough: (lineEnd: number): FootnoteAdmission => {
        const count = eventsThrough[Math.min(lineEnd, total)]!;
        planEvents(count);
        if (plans.length < count) throw new Error('Footnote plan requested beyond its proven rejection frontier');
        const admitted = plans.slice(0, count);
        return Object.freeze({
          ids: Object.freeze(admitted.flatMap((admission) => admission.ids)),
          layouts: Object.freeze(admitted.flatMap((admission) => admission.layouts)),
          reservePt: cumulativePt[count]!,
          continuations: Object.freeze(admitted.flatMap((admission) => admission.continuations)),
        });
      },
    });
  };
  let previousParagraph: BodyParagraphSourceInput | null = null;
  let terminalDiagnostic: LayoutDiagnostic | null = null;
  type PageEntryResume = Readonly<{ entryIndex: number }> & (
    | Readonly<{ kind: 'entry' }>
    | Readonly<{ kind: 'paragraph'; cursor: ParagraphFragmentCursor | null; suppressBefore: boolean }>
    | Readonly<{ kind: 'table'; cursor: BodyTableContinuationCursor | undefined }>
  );
  let currentEntryResume: () => PageEntryResume = () => ({ kind: 'entry', entryIndex: 0 });
  let resumeEntry: PageEntryResume | undefined;
  let restartEntryIndex = 0;
  let pageAcceptances: string[] = [];
  let restorePageCheckpoint: () => PageEntryResume = () => { throw new Error('Missing page note checkpoint'); };
  let installModernNoteArea = () => {};
  const acceptPageNode = (...args: Parameters<typeof acceptNode>): BodyPaginationState => {
    if (pageLocalNotes) pageAcceptances.push(bodyOccurrenceKey(args[2], acquisitionLocation(args[0]).flowDomainId, args[4]));
    return acceptNode(...args);
  };
  capturePageCheckpoint = (nextEntryIndex) => {
    if (!pageLocalNotes) return;
    const current = currentEntryResume();
    const resume: PageEntryResume = current.entryIndex === nextEntryIndex ? current
      : { kind: 'entry', entryIndex: nextEntryIndex };
    const savedState = state;
    const savedPrevious = previousParagraph;
    const savedPendingFrame = pendingFrameRun;
    const savedAutomatic = automaticPageStartEntryIndex;
    const savedHard = standaloneHardPageStartEntryIndex;
    const savedContinuousNotePage = continuousSectionNotePage;
    const allocationCount = allocations.length;
    const anchorCount = anchorInputs.length;
    const frameCount = frameObservations.length;
    const frameAnchorCount = frameAnchorObservations.length;
    const savedPendingNotes = [...pendingFootnoteContinuations];
    pageAcceptances = [];
    pageBandKeys = new Set();
    // Registries below own only their active page; the following prescan
    // reconstructs them through the canonical acquisition session.
    registeredAnchorPage = -1; registeredAnchorKeys.clear();
    deferralPage = -1; appliedDeferralKeys.clear();
    restorePageCheckpoint = () => {
      pageAcceptances.forEach(id => acceptedOccurrenceIds.delete(id));
      pageAcceptances.length = 0;
      allocations.length = allocationCount;
      anchorInputs.length = anchorCount;
      serializedAnchorInputs.length = anchorCount;
      frameObservations.length = frameCount;
      frameAnchorObservations.length = frameAnchorCount;
      const page = savedState.flow.pageIndex;
      footnoteIdsByPage.delete(page); footnoteReserveByPage.delete(page);
      footnoteLayoutsByPage.delete(page); footnoteHostShiftsByPage.delete(page);
      columnBandReservesByPage.delete(page); incomingTailPages.delete(page);
      pageBandKeys.forEach(band => { footnoteIdsByBand.delete(band); footnoteReserveByBand.delete(band); });
      pageBandKeys.clear();
      pendingFootnoteContinuations.clear();
      savedPendingNotes.forEach(([id, cursor]) => pendingFootnoteContinuations.set(id, cursor));
      previousParagraph = savedPrevious;
      pendingFrameRun = savedPendingFrame;
      automaticPageStartEntryIndex = savedAutomatic;
      standaloneHardPageStartEntryIndex = savedHard;
      continuousSectionNotePage = savedContinuousNotePage;
      registeredAnchorPage = -1; registeredAnchorKeys.clear();
      deferralPage = -1; appliedDeferralKeys.clear();
      terminalDiagnostic = null;
      state = savedState;
      return resume;
    };
  };
  installModernNoteArea = () => {
    if (!noteController) return;
    const area = noteController.area;
    const pageIndex = state.flow.pageIndex;
    const layouts = area.chunks.map(chunk => Object.freeze({ ...chunk.layout,
      pageArea: Object.freeze({ heightPt: area.heightPt, offsetPt: chunk.offsetPt }),
    }));
    footnoteLayoutsByPage.set(pageIndex, layouts);
    footnoteIdsByPage.set(pageIndex, new Set(noteController.resetActual()));
    footnoteReserveByPage.set(pageIndex, area.heightPt);
    retainBandReserve(pageFootnoteBandId(pageIndex), area.heightPt);
    state = setBodyFootnoteReserve(state, area.heightPt);
    const shifts = new Map<string, number>();
    area.chunks.forEach(chunk => {
      if (chunk.hostShiftPt !== 0) shifts.set(noteBandKey(chunk.layout.flowDomainId,
        chunk.layout.source.storyInstance), chunk.hostShiftPt);
    });
    if (shifts.size > 0) footnoteHostShiftsByPage.set(pageIndex, shifts);
    pendingFootnoteContinuations.clear();
    area.remaining.forEach(source => {
      if (!source.cursor) throw new Error('A page note plan omitted a required note head');
      pendingFootnoteContinuations.set(source.id, source.cursor);
    });
    if (area.chunks.some(chunk => incomingTailSourceIds.has(chunk.layout.source.storyInstance)
      && chunk.layout.story.blocks.some(block => block.kind === 'paragraph' && block.lines.length > 0))) {
      incomingTailPages.add(pageIndex);
    }
  };
  let incomingTailSourceIds = new Set<string>();
  const createModernController = (sources: readonly PageNoteSource[]) => {
    const location = mode15NoteOwner ?? acquisitionLocation(state);
    const pageIndex = state.flow.pageIndex;
    const section = location.section;
    const ownerRegion = state.pages.at(-1)!.accumulator.sectionRegions.find(region =>
      region.section === section) ?? activeRegion(state);
    const columns = ownerRegion.columnFlowDirection === 'rtl' ? [...section.columns].reverse() : section.columns;
    const maximumPt = state.flow.pageContentEndBlockPt - state.flow.pageContentStartBlockPt;
    const band = `${pageFootnoteBandId(pageIndex)}:band:${ownerRegion.id}:area`;
    const continuingPage = sources[0]?.cursor != null;
    incomingTailSourceIds = new Set(sources.map(source => source.id));
    return new PageNoteController(sources, maximumPt, (queue, capacityPt, balanced) => {
      let separatorAdvancePt: number | undefined;
      const acquire = (source: PageNoteSource, columnIndex: number, firstInColumn: boolean, availablePt: number) => {
        const column = columns[columnIndex]!;
        const noteLocation = Object.freeze({ ...location, section,
          availableBounds: Object.freeze({ ...location.availableBounds, xPt: column.xPt, widthPt: column.wPt }),
        });
        const result = acquireNoteColumnPartition(session, {
          kind: 'footnote', referenceIds: Object.freeze([source.id]), pageIndex, section,
          firstOnPage: firstInColumn, continuing: continuingPage,
          ...(columnIndex > 0 && firstInColumn && separatorAdvancePt !== undefined
            ? { blankLeadingAdvancePt: separatorAdvancePt } : {}),
          ...(reserves[pageIndex]?.footnoteTopsPt ? { plannedTopsPt: reserves[pageIndex]!.footnoteTopsPt } : {}),
          container: { id: `${band}:${columnIndex}`, kind: 'footnote', capacity: 'unbounded', bounds: {
            xPt: column.xPt, yPt: 0, widthPt: column.wPt, heightPt: section.geometry.pageHeight,
          } },
        }, source.cursor, availablePt, partitionFootnote,
        columnIndex === columns.length - 1
          ? continuationNoticeFor(noteLocation, column.wPt, false, true) : undefined,
        undefined, note => {
          if (columnIndex === 0 && firstInColumn) separatorAdvancePt = note.leading?.advancePt ?? note.story.flowBounds.yPt - note.flowBounds.yPt;
        });
        return source.allowPartial === false && result?.nextCursor ? null : result;
      };
      const request = { sources: queue, columnCount: columns.length, capacityPt, acquire };
      return balanced ? balancePageNoteArea(request) : flowPageNoteArea(request);
    }, services.allowFootnoteContinuation === true);
  };
  modernFootnoteAdmission = (ids, mandatoryUnit, allowPartial) => {
    if (!pageLocalNotes || ids.length === 0) return undefined;
    mode15NoteOwner ??= acquisitionLocation(state);
    // One-column owners keep the established reference-line permission for
    // physical continuation. Whole-note queries resolve that permission before
    // creating a plan; a later section's columns never replace this owner.
    if (mode15NoteOwner.section.columns.length === 1 && allowPartial === undefined) return undefined;
    noteController ??= createModernController(Object.freeze([]));
    return noteController.discover(ids, mandatoryUnit, mode15NoteOwner.section.columns.length === 1 ? allowPartial : undefined);
  };
  openModernIncomingNotes = () => {
    if (!pageLocalNotes || acquisitionLocation(state).section.columns.length < 2) return false;
    const sources = Object.freeze([...pendingFootnoteContinuations]
      .map(([id, cursor]) => Object.freeze({ id, cursor })));
    mode15NoteOwner = acquisitionLocation(state);
    noteController = createModernController(sources);
    if (noteController.area.chunks.length === 0) {
      throw new NoteCapacityExceededError('footnote', state.flow.pageIndex, pageFootnoteBandId(state.flow.pageIndex));
    }
    installModernNoteArea();
    return true;
  };
  resetNotePage = () => { noteController = undefined; mode15NoteOwner = undefined; incomingTailSourceIds = new Set(); };
  beforeNotePageClose = () => noteController?.beforePageClose();
  capturePageCheckpoint(0);
  yield* prescanPageAnchors(state, 0);

  for (;;) {
    try {
  bodyEntries: for (let entryIndex = restartEntryIndex; entryIndex < input.sequence.length; entryIndex += 1) {
    const resumed = resumeEntry?.entryIndex === entryIndex ? resumeEntry : undefined;
    resumeEntry = undefined;
    currentEntryResume = () => ({ kind: 'entry', entryIndex });
    // Suspension point. `state.pages.length` is the committed page count, which
    // is what a progressive driver reports; a synchronous driver discards it.
    observer?.onPageDrafts?.(state.pages);
    yield state.pages.length;
    // Resume after the scheduler has had the chance to yield to the host. A
    // first publication must never run in the same uninterrupted task as the
    // layout work that produced it, otherwise the browser cannot paint it.
    if (observer?.shouldPublish(state.pages.length)) {
      observer.publish(paginationPassResult(
        state,
        owners,
        session,
        allocations,
        footnoteReserveByPage,
        footnoteLayoutsByPage,
        footnoteHostShiftsByPage,
        terminalDiagnostic,
        [],
        [],
        frameObservations,
      ), entryIndex);
    }
    const entry = input.sequence[entryIndex]!;
    if (entry.kind === 'consume-source') {
      continue;
    }
    if (entry.kind === 'authored-break') {
      previousParagraph = null;
      const pageIndexBeforeBreak = state.flow.pageIndex;
      yield* commitTransition(
        applyAuthoredBreak(state.flow, entry.break, entry.parity),
        entryIndex + 1,
      );
      standaloneHardPageStartEntryIndex = entry.break === 'page'
        && entry.origin === 'authored'
        && entry.parity === undefined
        && entry.sameSourceParagraphAsPrevious !== true
        && state.flow.pageIndex !== pageIndexBeforeBreak
        ? entryIndex + 1
        : null;
      continue;
    }
    if (entry.kind === 'begin-section') {
      previousParagraph = null;
      const currentWritingMode = sectionWritingMode(activeRegion(state).section);
      const incomingWritingMode = sectionWritingMode(
        sectionContextForPage(entry.section, state.flow.pageIndex),
      );
      const currentPhysical = uprightPhysicalExtent({
        widthPt: activeRegion(state).section.geometry.pageWidth,
        heightPt: activeRegion(state).section.geometry.pageHeight,
      }, currentWritingMode);
      const incomingContext = sectionContextForPage(entry.section, state.flow.pageIndex);
      const incomingPhysical = uprightPhysicalExtent({
        widthPt: incomingContext.geometry.pageWidth,
        heightPt: incomingContext.geometry.pageHeight,
      }, incomingWritingMode);
      // §17.6.20 changes the logical page frame; one physical page cannot own
      // section regions whose logical axes use different writing modes.
      const effectiveStartType = entry.section.startType === 'continuous'
        && (currentWritingMode !== incomingWritingMode
          || currentPhysical.widthPt !== incomingPhysical.widthPt
          || currentPhysical.heightPt !== incomingPhysical.heightPt)
        ? 'nextPage'
        : entry.section.startType;
      const incomingInterval = pageBodyInterval(
        entry.section,
        state.flow.pageIndex,
        reserves[state.flow.pageIndex] ?? { top: 0, bottom: 0 },
      );
      const notesBeforeBoundary = hasFootnoteReferenceOnPage(state.flow.pageIndex);
      // WORD_CONTINUOUS_SECTION_FOOTNOTE_PAGE separates the mode15 observation
      // from the normative WW8 compatibility flag and its later-reference limit.
      const mayShareFootnotePage = wordContinuousSectionMayShareFootnotePage(
        input.compatibilityMode, input.footnoteLayoutLikeWW8, pageHasColumnFootnoteBands(state.flow.pageIndex));
      const boundaryPage = state.flow.pageIndex;
      try {
        yield* commitTransition(
          beginSection(
            state.flow,
            flowSection(entry.section, state.flow.pageIndex),
            effectiveStartType,
            {
              hasFootnoteReferenceOnCurrentPage: notesBeforeBoundary && !mayShareFootnotePage,
              incomingPageContentStartBlockPt: incomingInterval.blockStartPt,
              incomingPageContentEndBlockPt: incomingInterval.blockEndPt,
            },
          ),
          entryIndex + 1,
        );
        if (notesBeforeBoundary && mayShareFootnotePage && input.compatibilityMode !== 15
          && effectiveStartType === 'continuous' && state.flow.pageIndex === boundaryPage) {
          continuousSectionNotePage = boundaryPage;
        }
      } catch (error) {
        // beginSection rejects this authored transition before mutating the
        // immutable flow state. Once at least one physical page is complete,
        // retain only those committed pages and expose the omitted suffix as a
        // structured diagnostic. First-page failures and all invariant errors
        // remain fatal because no safe checkpoint exists for them.
        if (!(error instanceof UnsupportedPageFlowTransitionError)
          || state.flow.pageIndex === 0) {
          throw error;
        }
        const failedPageIndex = state.flow.pageIndex;
        state = Object.freeze({
          ...state,
          pages: Object.freeze(state.pages.filter((draft) => (
            draft.accumulator.pageIndex < failedPageIndex
          ))),
        });
        terminalDiagnostic = Object.freeze({
          code: 'UNSUPPORTED_FEATURE',
          severity: 'error',
          source: entry.source,
          message: 'Document layout stopped after the last complete page because '
            + `a nextColumn section could not be placed safely (${error.reason})`,
        });
        break bodyEntries;
      }
      continue;
    }
    const block = entry.kind === 'adjacent-table-group' ? entry : entry.block;
    if (block.kind === 'paragraph') {
      if (block.continuousSectionRole === 'collapse-mark') {
        continue;
      }
      if (block.pageBreakBefore && resumed?.kind !== 'paragraph') {
        yield* commitTransition(
          applyAuthoredBreak(state.flow, 'pageBreakBefore'),
          entryIndex,
        );
      }
      const previousAfterPt = previousParagraph?.spaceAfterPt ?? 0;
      const spacing = resumed?.kind === 'paragraph'
        ? { suppressBefore: resumed.suppressBefore, overlap: 0 }
        : paragraphGapAdjustment(previousParagraph, block, previousAfterPt,
          block.continuousSectionRole === 'suppress-before' ? 0 : block.spaceBeforePt);
      const spacingOverlap = block.continuousSectionRole === 'drop-previous-after'
        ? previousAfterPt
        : spacing.overlap;
      // WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: the preceding in-flow cursor,
      // including its space after, anchors every frame member. Only the next
      // in-flow paragraph consumes that after/before overlap, once.
      if (resumed?.kind !== 'paragraph' && !block.textFrameMember && spacingOverlap > 0) {
        state = Object.freeze({
          ...state,
          flow: Object.freeze({
            ...state.flow,
            cursorBlockPt: Math.max(
              state.flow.regionStartBlockPt,
              state.flow.cursorBlockPt - spacingOverlap,
            ),
          }),
        });
      }
      let cursor: ParagraphFragmentCursor | null = resumed?.kind === 'paragraph'
        ? resumed.cursor : Object.freeze({ boundary: null });
      currentEntryResume = () => ({ kind: 'paragraph', entryIndex, cursor, suppressBefore: spacing.suppressBefore });
      // A text-anchored frame takes no flow space, so the next in-flow
      // paragraph collapses its spacing against the text before the frames
      // (WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH).
      while (cursor) {
        const fragmentStartKey = paragraphFragmentStartKey(cursor);
        let location = acquisitionLocation(state);
        const acquired = session.measureParagraph({
          input: block,
          location,
          availableInlineExtentPt: location.availableBounds.widthPt,
          suppressSpaceBefore: cursor.boundary !== null
            || block.continuousSectionRole === 'suppress-before'
            || spacing.suppressBefore
            || (
              cursor.boundary === null
              // Windows trr3 C: automatic overflow suppresses before for
              // empty marks as well as text. Objects remain outside the controls.
              && (block.inkless === true || block.onlyTextualContent === true)
              && !state.flow.pageHasContent
              && automaticPageStartEntryIndex === entryIndex
            )
            || (
              // Keep-with-next relocation owns a separate, content-independent
              // leading-spacing rule for the complete group.
              cursor.boundary === null
              && block.keepNext
              && block.inkless !== true
              && !state.flow.pageHasContent
              && automaticPageStartEntryIndex === entryIndex
            )
            || (
              cursor.boundary === null
              && block.pageBreakBefore
              && input.compatibilityMode === 15
              && !state.flow.pageHasContent
            )
            || (
              cursor.boundary === null
              && !block.pageBreakBefore
              && !state.flow.pageHasContent
              && standaloneHardPageStartEntryIndex === entryIndex
            ),
          continuation: cursor,
        });
        if (acquired.placement) {
          const notes = footnoteAdmission(
            acquired.layout,
            location.availableBounds.widthPt,
            acquired.retainedFootnoteReferenceIds,
          );
          const relocationExtentPt = acquired.relocationBlockExtentPt;
          const admissionChargePt = acquired.placement.sectionFlowOwnership === 'page'
            ? notes.reservePt
            : (relocationExtentPt ?? acquired.blockExtentPt) + notes.reservePt;
          const freshExtentPt = freshPageExtent(state);
          // The footnote band is physical-page global; spare room in the active
          // column cannot authorize a reserve that clips a deeper prior column.
          const reserveInvadesCommittedPageContent =
            footnoteReserveInvadesCommittedPageContent(notes.reservePt);
          assertFreshPageFootnoteAdmission(
            notes.reservePt,
            admissionChargePt,
            freshExtentPt,
          );
          if (
            (
              admissionChargePt > location.availableBounds.heightPt
              || reserveInvadesCommittedPageContent
            )
            && admissionChargePt <= freshExtentPt
            && pageCanGainSpaceByRelocating()
          ) {
            yield* commitTransition(
              reserveInvadesCommittedPageContent
                ? advanceToPage(state.flow, state.flow.section, 'overflow')
                : advanceColumnOrPage(state.flow, 'overflow'),
              entryIndex,
            );
            continue;
          }
          if (acquired.frameCarry && frameRunDeferred(acquired.frameCarry.runKey)) {
            yield* commitTransition(advanceColumnOrPage(state.flow, 'overflow'), entryIndex);
            continue;
          }
          if (acquired.frameCarry) {
            const frameDomainId = acquisitionLocation(state).flowDomainId;
            if (pendingFrameRun?.runKey !== acquired.frameCarry.runKey) {
              pendingFrameRun = Object.freeze({
                runKey: acquired.frameCarry.runKey,
                framesPageIndex: state.flow.pageIndex,
                framesFlowDomainId: frameDomainId,
                framesAfterContent: state.flow.pageHasContent,
              });
            }
            frameObservations.push(Object.freeze({
              nodeId: bodyOccurrenceKey(block.source, frameDomainId, fragmentStartKey),
              pageIndex: state.flow.pageIndex,
              flowDomainId: frameDomainId,
              meetsEarlierContent: frameMeetsEarlierFlowContent(
                state, frameDomainId, acquired.frameCarry,
              ),
              carry: acquired.frameCarry,
            }));
          }
          // A placed frame still paints its retained references despite a zero flow
          // charge, so note ownership is committed with the accepted occurrence.
          state = acceptPageNode(
            state,
            acquired.layout,
            block.source,
            acquired.blockExtentPt,
            fragmentStartKey,
            acceptedOccurrenceIds,
            allocations,
            acquired.placement,
          );
          commitFootnotes(notes.ids, notes.layouts);
          if (acquired.flowRegistryDelta) {
            session.commitFlowRegistryDelta(acquired.flowRegistryDelta);
          }
          cursor = null;
          session.moveAcquisitionCursor(acquisitionLocation(state));
          continue;
        }
        // A same-page column reaching the physical bottom has the same edge;
        // a balanced/shortened column, footer or existing note band does not.
        const region = activeRegion(state);
        const physicalBodyBottomPt = region.section.geometry.pageHeight
          - Math.abs(region.section.geometry.marginBottom);
        const unreservedBodyBottomPt = region.writingMode === 'horizontal-tb'
          && activeBlockEndPt(state) === physicalBodyBottomPt
          && (reserves[state.flow.pageIndex]?.bottom ?? 0) === 0
          && state.footnoteReservePt === 0
          ? physicalBodyBottomPt : undefined;
        if (cursor.boundary === null && block.keepNext && state.flow.pageHasContent) {
          let keepSetExtentPt = acquired.blockExtentPt;
          const keepSetReferenceIds = new Set(footnoteIdsInRetainedSlice(acquired.layout));
          let terminalGridParagraph: ParagraphLayout | undefined;
          let terminalLeadExtentPt = 0;
          let hasTerminalBlock = false;
          let keepSetBlockedByDeferral = false;
          let bridgeSuccessor = wordEmptyKeepNextBridgesSuccessor({
            keepNext: block.keepNext,
            inkless: block.inkless === true,
            undecoratedMark: isUndecoratedInklessMark(acquired.layout),
          });
          for (let nextIndex = entryIndex + 1; nextIndex < input.sequence.length; nextIndex += 1) {
            const nextEntry = input.sequence[nextIndex]!;
            if (nextEntry.kind === 'consume-source') continue;
            if (nextEntry.kind === 'authored-break' || nextEntry.kind === 'begin-section') break;
            const nextBlock = nextEntry.kind === 'adjacent-table-group'
              ? nextEntry
              : nextEntry.block;
            if (nextBlock.kind === 'paragraph' && nextBlock.pageBreakBefore) break;
            const following = session.measureFollowingBlock({
              input: nextBlock,
              location,
              availableInlineExtentPt: location.availableBounds.widthPt,
            });
            const continues = nextBlock.kind === 'paragraph'
              && (nextBlock.keepNext || bridgeSuccessor);
            bridgeSuccessor = false;
            // A proven anchor-line deferral on a line this keep set must hold
            // ends the page above it, like an overflow
            // (word-page-anchor-line-deferral).
            const keptLines = continues
              ? following.pageOwnedAnchorKeysByLine
              : following.pageOwnedAnchorKeysByLine?.slice(0, 1);
            if (keptLines?.some((keys) => anchorLineDeferred(keys))) {
              keepSetBlockedByDeferral = true;
            }
            keepSetExtentPt += continues
              ? following.fullExtentPt
              : following.leadContentExtentPt;
            const referenceIds = continues
              ? following.fullFootnoteReferenceIds
              : following.leadFootnoteReferenceIds;
            referenceIds?.forEach((id) => keepSetReferenceIds.add(id));
            if (!continues) {
              terminalGridParagraph = following.leadGridParagraph;
              terminalLeadExtentPt = following.leadContentExtentPt;
            }
            if (!continues || keepSetBlockedByDeferral) {
              hasTerminalBlock = true;
              break;
            }
            // The existing admission branch below can relocate only a
            // complete keep set whose total charge fits a fresh page. Once
            // the running block extent alone exceeds that bound, the branch
            // cannot be taken and measuring the rest of the chain is wasted
            // work. Long keep-with-next chains otherwise re-measure the shared
            // tail once per member, which is effectively unbounded on
            // pathological documents.
            if (keepSetExtentPt > freshPageExtent(state)) break;
          }
          const keepSetReservePt = footnoteAdmissionForIds(
            [...keepSetReferenceIds],
            location.availableBounds.widthPt,
          ).reservePt;
          const terminalAdmissionPt = terminalGridParagraph
            ? wordGridLineAdmissionExtentPt(terminalGridParagraph, terminalLeadExtentPt,
              unreservedBodyBottomPt, keepSetReservePt, keepSetExtentPt - terminalLeadExtentPt)
            : terminalLeadExtentPt;
          const keepSetAdmissionPt = keepSetExtentPt
            - (terminalLeadExtentPt - terminalAdmissionPt) + keepSetReservePt;
          if (
            hasTerminalBlock
            && (keepSetAdmissionPt > location.availableBounds.heightPt || keepSetBlockedByDeferral)
            && keepSetAdmissionPt <= freshPageExtent(state)
          ) {
            yield* commitTransition(
              advanceColumnOrPage(state.flow, 'overflow'),
              entryIndex,
              true,
            );
            continue;
          }
        }
        const nextEntry = input.sequence[entryIndex + 1];
        const afterNext = input.sequence[entryIndex + 2];
        const isImmediatelyPreBreakAnchor = nextEntry?.kind === 'body-block'
          && nextEntry.block.kind === 'paragraph'
          && afterNext?.kind === 'authored-break'
          && afterNext.break === 'page'
          // A trailing hard break belongs to the anchor's own source paragraph.
          // It governs what follows that paragraph; it is not evidence that the
          // preceding inline resource and anchor form a movable keep group.
          && afterNext.sameSourceParagraphAsPrevious !== true;
        // `word-pre-break-inline-drawing-group`
        const hasInlineDrawingResource = wordPreBreakInlineDrawingResource(acquired.layout);
        if (
          cursor.boundary === null
          && hasInlineDrawingResource
          && isImmediatelyPreBreakAnchor
          && state.flow.pageHasContent
        ) {
          const anchorLocation = locationAfter(location, acquired.blockExtentPt);
          const anchor = session.measureParagraph({
            input: nextEntry.block,
            location: anchorLocation,
            availableInlineExtentPt: anchorLocation.availableBounds.widthPt,
            suppressSpaceBefore: false,
            continuation: Object.freeze({ boundary: null }),
          });
          session.moveAcquisitionCursor(location);
          const anchorExtentPt = wordPreBreakHostAnchorExtentPt(
            anchor.layout,
            anchorLocation.cursorPt.yPt,
          );
          if (anchorExtentPt !== null) {
            const groupExtentPt = acquired.blockExtentPt + anchorExtentPt;
            if (
              groupExtentPt > location.availableBounds.heightPt
              && groupExtentPt <= freshPageExtent(state)
            ) {
              yield* commitTransition(
                advanceColumnOrPage(state.flow, 'overflow'),
                entryIndex,
              );
              continue;
            }
          }
        }
        const followingEntry = input.sequence[entryIndex + 1];
        const followedByHardPageBreak = followingEntry?.kind === 'authored-break'
          && followingEntry.break === 'page';
        if (
          cursor.boundary === null
          && followedByHardPageBreak
        ) {
          // `word-pre-break-anchor-paragraph`
          const neutral = wordFlowNeutralPreBreakAnchorParagraph(acquired.layout);
          if (neutral !== null) {
            state = acceptPageNode(
              state,
              neutral,
              block.source,
              0,
              fragmentStartKey,
              acceptedOccurrenceIds,
              allocations,
            );
            if (acquired.flowRegistryDelta) {
              session.commitFlowRegistryDelta(acquired.flowRegistryDelta);
            }
            cursor = null;
            session.moveAcquisitionCursor(acquisitionLocation(state));
            continue;
          }
        }
        // The empty-mark allowance can apply only to an undecorated inkless
        // mark that starts here without keepNext; any other fragment ignores
        // the reserve, so its whole-note acquisition is not performed (it
        // would otherwise touch every note of a long paragraph per attempt).
        const markReserveMatters = cursor.boundary === null && block.inkless === true
          && !block.keepNext && isUndecoratedInklessMark(acquired.layout);
        const markReservePt = markReserveMatters
          ? footnoteAdmission(acquired.layout, location.availableBounds.widthPt).reservePt
          : 0;
        const pageBottomIsUnreserved = (reserves[state.flow.pageIndex]?.bottom ?? 0) === 0
          && state.footnoteReservePt === 0;
        const physicalRegionBottomIsActive = activeBlockEndPt(state) === activeRegion(state).blockEndPt;
        const followsNextPageSectionBoundary = followingEntry?.kind === 'begin-section'
          && followingEntry.section.startType === 'nextPage';
        // Compatibility-owned physical-edge empty-mark admission.
        const trailingMarkAdmissionAllowancePt =
          wordTrailingEmptyMarkAdmissionAllowancePt({
            hasContinuationBoundary: cursor.boundary !== null,
            inkless: block.inkless === true,
            undecorated: isUndecoratedInklessMark(acquired.layout),
            keepNext: block.keepNext,
            markReservePt,
            pageBottomIsUnreserved,
            physicalRegionBottomIsActive,
            hasFollowingInk: hasFollowingInkContent(input, entryIndex + 1),
            followsNextPageSectionBoundary,
            markExtentPt: acquired.blockExtentPt,
            markBelowBaselinePt: acquired.markBelowBaselinePt ?? 0,
            markOnLineGrid: acquired.markOnLineGrid === true,
            markGridAfterLeadingPt: acquired.markGridAfterLeadingPt,
          });
        // A proven anchor-line deferral ends this page just above that line.
        const anchorLineBreak = deferredAnchorLineIndex(acquired.layout);
        // Same fresh-region rule as framed paragraphs: a page whose band
        // holds carried note lines is not fresh (pageCanGainSpaceByRelocating).
        const canRelocateParagraph = pageCanGainSpaceByRelocating();
        // Continuation mode (the public default) reserves each referenced note
        // at its reference line; later lines are admitted only beside those
        // committed plans. With continuation off every note stays whole.
        const notePlans = services.allowFootnoteContinuation === true
          ? referenceLineFootnotePlans(acquired.layout, {
              canRelocate: canRelocateParagraph,
              indivisible: acquired.fragmentation.kind === 'indivisible',
              continuesFromPrevious: cursor.boundary !== null,
              entryIndex,
              keepLines: cursor.boundary === null && block.keepLines,
              widowControl: block.widowControl,
              lineEndLimit: anchorLineBreak,
              admissionLimitPt: location.availableBounds.heightPt + trailingMarkAdmissionAllowancePt,
              authoredSpaceAfterPt: block.spaceAfterPt,
              writingMode: activeRegion(state).writingMode,
            })
          : null;
        const selected = selectParagraphFragment(
          acquired.layout,
          cursor,
          acquired.fragmentation,
          location.availableBounds.heightPt + trailingMarkAdmissionAllowancePt,
          freshPageExtent(state),
          canRelocateParagraph,
          {
            keepLines: block.keepLines,
            widowControl: block.widowControl,
            authoredSpaceAfterPt: block.spaceAfterPt,
            ...(unreservedBodyBottomPt === undefined ? {} : { unreservedBodyBottomPt }),
            writingMode: activeRegion(state).writingMode,
            ...(anchorLineBreak === undefined ? {} : { lineEndLimit: anchorLineBreak }),
          },
          (fragment) => notePlans
            ? notePlans.reservePt(fragment.lines.length)
            : footnoteAdmission(fragment, location.availableBounds.widthPt).reservePt,
          acquired.uniformRubyAdvancePt,
          (reservePt) => !footnoteReserveInvadesCommittedPageContent(reservePt),
        );
        if (selected.requiresFreshFlowRegion) {
          yield* commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
            true,
          );
          continue;
        }
        if (!selected.fragment) throw new Error('Paragraph acquisition made no progress');
        // Commit exactly the plan objects charged during selection, taken
        // against the pre-acceptance page state; plans of rejected later
        // reference lines are never committed.
        const plannedNotes = notePlans?.admissionThrough(selected.fragment.lines.length);
        if (pendingFrameRun && cursor.boundary === null) {
          const at = acquisitionLocation(state);
          frameAnchorObservations.push(Object.freeze({
            ...pendingFrameRun,
            anchorPageIndex: at.pageIndex,
            anchorFlowDomainId: at.flowDomainId,
          }));
          pendingFrameRun = null;
        }
        state = acceptPageNode(
          state,
          selected.fragment,
          block.source,
          Math.min(selected.admittedBlockExtentPt, location.availableBounds.heightPt),
          fragmentStartKey,
          acceptedOccurrenceIds,
          allocations,
          acquired.placement,
        );
        const notes = plannedNotes
          ?? footnoteAdmission(selected.fragment, location.availableBounds.widthPt);
        // The fatal guard re-checks the fit the selector admitted, with the
        // same extent (authored trailing space after a completing fragment is
        // not charged), so it rejects only notes that cannot fit any fresh
        // page beside that fragment. Retained advance still drives the cursor.
        assertFreshPageFootnoteAdmission(
          notes.reservePt,
          selected.admissionExtentPt + notes.reservePt,
          freshPageExtent(state),
        );
        commitFootnotes(notes.ids, notes.layouts, notes.continuations);
        if (acquired.flowRegistryDelta) {
          const acceptedDelta = paragraphFlowRegistryDeltaForAcceptedFragment(
            acquired.flowRegistryDelta,
            selected.fragment,
          );
          if (acceptedDelta) session.commitFlowRegistryDelta(acceptedDelta);
        }
        cursor = selected.nextCursor;
        if (cursor) {
          yield* commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
        }
        location = acquisitionLocation(state);
        session.moveAcquisitionCursor(location);
      }
      if (!block.textFrameMember) previousParagraph = block;
    } else {
      previousParagraph = null;
      // Only a paragraph anchors a frame run (ECMA-376 §17.3.1.11).
      pendingFrameRun = null;
      currentEntryResume = () => ({ kind: 'table', entryIndex, cursor: undefined });
      if (block.kind === 'table' && resumed?.kind !== 'table') {
        const tableKey = `table:${sourceKey(block.source)}`;
        const minimumPage = minimumTablePageBySource?.get(tableKey);
        if (block.pageOwnedFloatingTable === true || minimumPage !== undefined) {
          recordAnchorInput(Object.freeze({
            kind: 'page-owned-table',
            pageIndex: state.flow.pageIndex,
            key: tableKey,
            floor: minimumPage,
          }));
        }
        if (minimumPage !== undefined) {
          // §17.4.57 gives the page-owned table a fixed position and a minimum
          // distance from adjacent text. Word controls show that a table whose
          // own exclusion sends its preceding text past candidate page p is
          // instead tried on p+1. Do not jump to the observed source page:
          // intermediate pages may still admit the table. This floor belongs
          // only to this exact convergence run.
          while (state.flow.pageIndex < minimumPage) {
            yield* commitTransition(
              advanceToPage(state.flow, state.flow.section, 'overflow'),
              entryIndex,
            );
          }
        }
      }
      let cursor: BodyTableContinuationCursor | undefined = resumed?.kind === 'table' ? resumed.cursor : undefined;
      const tableResume = (): PageEntryResume => ({ kind: 'table', entryIndex, cursor });
      currentEntryResume = tableResume;
      // An owner segment pending at the end of a region starts the next one
      // as a fresh region (table-owner-runs.ts): every transition that leaves
      // the region with a same-region entry pending passes through here.
      const enteringFreshRegion = (
        next: BodyTableContinuationCursor | undefined,
      ): BodyTableContinuationCursor | undefined => next?.ownerSegmentEntry === 'same-region'
        ? Object.freeze({ ...next, ownerSegmentEntry: 'fresh-region' as const })
        : next;
      let complete = false;
      let unwrappedLocation: BodyAcquisitionLocation | undefined;
      while (!complete) {
        const fragmentStartKey = tableFragmentStartKey(cursor);
        const location = acquisitionLocation(state);
        if (!unwrappedLocation || unwrappedLocation.flowDomainId !== location.flowDomainId
          || unwrappedLocation.pageIndex !== location.pageIndex
          || unwrappedLocation.columnIndex !== location.columnIndex) {
          unwrappedLocation = location;
        }
        const initialLocation = unwrappedLocation;
        const requestAt = (availableBlockExtentPt: number) => session.measureTable({
            input: block,
            location,
            unwrappedLocation: initialLocation,
            availableInlineExtentPt: location.availableBounds.widthPt,
            availableBlockExtentPt,
            freshPageBlockExtentPt: freshPageExtent(state),
            pageHasContent: state.flow.pageHasContent,
            ...(cursor ? { cursor } : {}),
          });
        let availableBlockExtentPt = location.availableBounds.heightPt;
        let acquired = requestAt(availableBlockExtentPt);
        if (acquired.retryAtBlockStartPt !== undefined) {
          if (!Number.isFinite(acquired.retryAtBlockStartPt)
            || acquired.retryAtBlockStartPt <= state.flow.cursorBlockPt) {
            throw new Error('Table repositioning must advance the block cursor');
          }
          state = Object.freeze({
            ...state,
            flow: Object.freeze({
              ...state.flow,
              cursorBlockPt: acquired.retryAtBlockStartPt,
            }),
          });
          session.moveAcquisitionCursor(acquisitionLocation(state));
          continue;
        }
        let notes = acquired.requiresFreshFlowRegion
          ? Object.freeze({
              ids: Object.freeze([]) as readonly string[],
              layouts: Object.freeze([]) as readonly NoteLayout[],
              reservePt: 0,
            })
          : footnoteAdmission(
              acquired.layout,
              location.availableBounds.widthPt,
            );
        // A zero-advance cell-owner host still occupies the host flow below the
        // cursor; its footnotes are admitted against that extent, as for a
        // placed paragraph frame. Other tables charge their block advance.
        const flowChargePt = (block: typeof acquired): number =>
          block.relocationBlockExtentPt ?? block.blockExtentPt;
        let lastFootnoteAdmission = Object.freeze({
          reservePt: notes.reservePt,
          chargePt: flowChargePt(acquired) + notes.reservePt,
        });
        const seenCandidates = new Set<string>();
        // The region extent and the fragment charge reconstruct the same
        // authored boundary through different arithmetic, so the comparison
        // must tolerate single-ULP drift exactly like the invariant layer
        // (invariants.ts atMostWithinFloatingPrecision, mirrored inline here
        // because the layout-boundary gate pins this module's invariants
        // import to exactly assertAndDeepFreezeDocumentLayout): a fragment
        // whose charge exceeds the region only within floating-point
        // precision is already placed, and with an unchanged footnote
        // reserve the retry request would be identical to the first — a
        // raw `>` here could never converge and would trip the fingerprint
        // guard below on floating-point dust.
        const fitsWithinFloatingPrecision = (chargePt: number, heightPt: number): boolean =>
          chargePt <= heightPt
          || chargePt - heightPt
            <= Number.EPSILON * Math.max(1, Math.abs(chargePt), Math.abs(heightPt));
        while (
          !acquired.requiresFreshFlowRegion
          && !fitsWithinFloatingPrecision(
            flowChargePt(acquired) + notes.reservePt,
            location.availableBounds.heightPt,
          )
        ) {
          const fingerprint = JSON.stringify({
            advancePt: acquired.blockExtentPt,
            chargePt: flowChargePt(acquired),
            nextCursor: acquired.nextCursor ?? null,
            noteIds: notes.ids,
            reservePt: notes.reservePt,
          });
          if (seenCandidates.has(fingerprint)) {
            assertFreshPageFootnoteAdmission(
              lastFootnoteAdmission.reservePt,
              lastFootnoteAdmission.chargePt,
              freshPageExtent(state),
            );
            throw new Error('Table footnote admission did not converge');
          }
          seenCandidates.add(fingerprint);
          availableBlockExtentPt = Math.max(
            0,
            location.availableBounds.heightPt - notes.reservePt,
          );
          acquired = requestAt(availableBlockExtentPt);
          notes = acquired.requiresFreshFlowRegion
            ? Object.freeze({
                ids: Object.freeze([]) as readonly string[],
                layouts: Object.freeze([]) as readonly NoteLayout[],
                reservePt: 0,
              })
            : footnoteAdmission(
                acquired.layout,
                location.availableBounds.widthPt,
              );
          if (!acquired.requiresFreshFlowRegion) {
            lastFootnoteAdmission = Object.freeze({
              reservePt: notes.reservePt,
              chargePt: flowChargePt(acquired) + notes.reservePt,
            });
          }
        }
        if (acquired.requiresFreshFlowRegion) {
          assertFreshPageFootnoteAdmission(
            lastFootnoteAdmission.reservePt,
            lastFootnoteAdmission.chargePt,
            freshPageExtent(state),
          );
          const rebasesFloatingTableOnFreshFrame = !state.flow.pageHasContent
            && acquired.nextCursor?.floatingContinuationFrame === 'fresh-text'
            && cursor?.floatingContinuationFrame === undefined;
          // A floating continuation frame or an owner-segment entry is state
          // the next region's request must carry; plain tables keep theirs.
          if (acquired.nextCursor?.floatingContinuationFrame !== undefined
            || acquired.nextCursor?.ownerSegmentEntry !== undefined) {
            cursor = acquired.nextCursor;
          }
          if (rebasesFloatingTableOnFreshFrame) continue;
          cursor = enteringFreshRegion(cursor);
          yield* commitTransition(
            acquired.requiresFreshPage
              ? advanceToPage(state.flow, state.flow.section, 'overflow')
              : advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
          continue;
        }
        if (
          footnoteReserveInvadesCommittedPageContent(notes.reservePt)
          && state.flow.pageHasContent
        ) {
          // The acquired table is already a coherent row fragment. A fresh
          // physical page preserves it; another same-page column cannot create
          // more room for the page-wide note band.
          cursor = enteringFreshRegion(cursor);
          yield* commitTransition(
            advanceToPage(state.flow, state.flow.section, 'overflow'),
            entryIndex,
          );
          continue;
        }
        state = acceptPageNode(
          state,
          acquired.layout,
          block.source,
          acquired.blockExtentPt,
          fragmentStartKey,
          acceptedOccurrenceIds,
          allocations,
          acquired.placement,
        );
        commitFootnotes(notes.ids, notes.layouts);
        if (acquired.flowRegistryDelta) {
          session.commitFlowRegistryDelta(bindTableFlowRegistryDeltaToAcceptedOccurrence(
            acquired.flowRegistryDelta,
            acquired.layout,
            bodyOccurrenceKey(block.source, location.flowDomainId, fragmentStartKey),
          ));
        }
        cursor = acquired.nextCursor ?? undefined;
        unwrappedLocation = undefined;
        complete = cursor === undefined;
        // WORD_OVER_PAGE_CELL_BREAK_OCCUPANCY:
        // clipped cell content counts a physical continuation page before a
        // following authored page break. The continuation page's normal-flow
        // cursor stays at its top: without a break, the next paragraph starts
        // there, not after the invisible remainder. The 500pt/800pt controls
        // bracketed a 648pt body band; use retained overflow height rather
        // than a fixed extra-page allowance.
        let hiddenOverflowPt = acquired.unpaintedOverflowPt ?? 0;
        if (hiddenOverflowPt > 0) {
          if (!Number.isFinite(hiddenOverflowPt)) throw new Error('Table overflow extent must be finite');
          // A body's usable extent may differ on first/even/odd pages because
          // header/footer reserves vary. For budget preflight, the full page
          // dimension is only an upper bound; actual charging below uses each
          // destination page's reserved body interval.
          const maxPageExtentPt = Math.max(
            state.flow.section.geometry.pageWidth,
            state.flow.section.geometry.pageHeight,
          );
          if (!Number.isFinite(maxPageExtentPt) || maxPageExtentPt <= 0) {
            throw new Error('Table overflow requires a finite positive page extent');
          }
          const minimumContinuationPages = Math.ceil(hiddenOverflowPt / maxPageExtentPt);
          if (state.flow.pageIndex + minimumContinuationPages >= MAX_BODY_LAYOUT_PAGES) {
            throw new Error(`Document page budget exceeded (${MAX_BODY_LAYOUT_PAGES} pages)`);
          }
          // Only the final hidden-occupancy page can later receive a body
          // reference and replay. Its source resumes after this accepted table,
          // or at the table's retained source cursor if it still continues.
          currentEntryResume = () => cursor ? tableResume() : { kind: 'entry', entryIndex: entryIndex + 1 };
          while (hiddenOverflowPt > 0) {
            const pageExtentPt = freshPageExtent(state);
            if (!Number.isFinite(pageExtentPt) || pageExtentPt <= 0) {
              throw new Error('Table overflow requires a finite positive page extent');
            }
            const hasMoreHiddenPages = hiddenOverflowPt > pageExtentPt;
            yield* commitTransition(
              advanceToPage(state.flow, state.flow.section, 'overflow'),
              entryIndex,
              false,
              hasMoreHiddenPages,
            );
            hiddenOverflowPt -= pageExtentPt;
          }
          cursor = enteringFreshRegion(cursor);
          currentEntryResume = tableResume;
        } else if (cursor && cursor.ownerSegmentEntry !== 'same-region') {
          yield* commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
        }
      }
    }
    session.moveAcquisitionCursor(acquisitionLocation(state));
  }
      if (terminalDiagnostic === null) beforeNotePageClose();
      break;
    } catch (error) {
      if (!(error instanceof PageNoteReplay)) throw error;
      resumeEntry = restorePageCheckpoint();
      restartEntryIndex = resumeEntry.entryIndex;
      session.resetPageAcquisition(acquisitionLocation(state));
      yield* prescanPageAnchors(state, restartEntryIndex);
      installModernNoteArea();
    }
  }
  currentEntryResume = () => ({ kind: 'entry', entryIndex: input.sequence.length });
  // ECMA-376 §17.11.22: document-end notes are flow following the main
  // story, not decoration of its terminal page. Paginate them here through
  // the same column/page transitions, before draining footnote tails, so
  // incoming footnotes constrain each destination and NUMPAGES/header/footer
  // convergence observes every note-only page. Composition only paints the
  // already acquired occurrences below; it must never catch capacity and
  // silently omit a note suffix.
  const endnoteLayoutsByPage = new Map<number, NoteLayout[]>();
  const retainedEndnoteIds = new Set(state.pages.flatMap(draft =>
    draft.accumulator.paint.flatMap(({ layer, node }) => layer === 'body'
      && (node.kind === 'paragraph' || node.kind === 'table')
      ? endnoteIdsInRetainedSlice(node) : [])));
  const endnoteIds = (input.endnoteIds ?? []).filter(id => retainedEndnoteIds.has(id));
  if (endnoteIds.length > 0 && !session.layoutNotes) {
    throw new Error('Document-end notes require a note-capable layout session');
  }
  // Continuation belongs to the endnote area, not just an individual note:
  // §17.11.4 applies also when the next band starts with a new note.
  let endnoteFlowStarted = false;
  const modernEndnotes = input.compatibilityMode === 15 && activeRegion(state).columns.length > 1;
  if (modernEndnotes) {
    let pending: readonly PageNoteSource[] = endnoteIds.map(id => Object.freeze({ id, cursor: null }));
    while (pending.length > 0) {
      const region = activeRegion(state);
      const section = region.section;
      const columns = region.columnFlowDirection === 'rtl' ? [...section.columns].reverse() : section.columns;
      const topPt = Math.max(state.flow.cursorBlockPt, state.flow.deepestColumnBlockPt);
      const capacityPt = Math.max(0, region.blockEndPt - state.footnoteReservePt - topPt);
      const band = `endnotes:page:${state.flow.pageIndex}:band:${region.id}`;
      const continuingPage = endnoteFlowStarted;
      // WORD_PAGE_WIDE_ENDNOTE_AREA: mode 15 flows a common area below
      // balanced body columns, using each authored column's own width.
      let separatorAdvancePt: number | undefined;
      const area = balancePageNoteArea({ sources: pending, columnCount: columns.length, capacityPt,
        acquire(source, columnIndex, firstInColumn, availablePt, offsetPt) {
          const column = columns[columnIndex]!;
          return acquireNoteColumnPartition(session, {
            kind: 'endnote', referenceIds: [source.id], pageIndex: state.flow.pageIndex,
            section, firstOnPage: firstInColumn, continuing: continuingPage,
            ...(columnIndex > 0 && firstInColumn && separatorAdvancePt !== undefined
              ? { blankLeadingAdvancePt: separatorAdvancePt } : {}),
            bandTranslationPt: Object.freeze({ xPt: 0, yPt: topPt + offsetPt }),
            container: { id: `${band}:${columnIndex}`, kind: 'endnote', capacity: 'unbounded', bounds: {
              xPt: column.xPt, yPt: 0, widthPt: column.wPt, heightPt: section.geometry.pageHeight,
            } },
          }, source.cursor, availablePt, partitionFootnote, undefined, undefined, note => {
            if (columnIndex === 0 && firstInColumn) separatorAdvancePt = note.leading?.advancePt ?? note.story.flowBounds.yPt - note.flowBounds.yPt;
          });
        },
      });
      if (area.chunks.length === 0) {
        if (topPt === state.flow.regionStartBlockPt && state.footnoteReservePt === 0) {
          throw new NoteCapacityExceededError('endnote', state.flow.pageIndex, band);
        }
      } else {
        endnoteLayoutsByPage.set(state.flow.pageIndex, area.chunks.map(chunk => translateNoteLayout(chunk.layout,
          { xPt: 0, yPt: topPt + chunk.offsetPt - chunk.layout.flowBounds.yPt })));
        pending = area.remaining;
        endnoteFlowStarted = true;
        state = Object.freeze({ ...state, flow: Object.freeze({ ...state.flow,
          pageHasContent: true, cursorBlockPt: topPt + area.heightPt,
          deepestColumnBlockPt: topPt + area.heightPt,
        }) });
      }
      if (pending.length > 0) {
        yield* commitTransition(advanceToPage(state.flow, state.flow.section, 'overflow'), input.sequence.length);
      }
      observer?.onPageDrafts?.(state.pages);
      yield state.pages.length;
    }
  }
  let endnoteBand = '';
  for (const id of modernEndnotes ? [] : endnoteIds) {
    let cursor: FootnoteCursor | null = null;
    let complete = false;
    while (!complete) {
      const location = acquisitionLocation(state);
      const band = `endnotes:page:${state.flow.pageIndex}`
        + (activeRegion(state).columns.length > 1 ? `:band:${activeRegion(state).id}:${state.flow.columnIndex}` : '');
      const partition = acquireNoteColumnPartition(session, {
        kind: 'endnote', referenceIds: Object.freeze([id]), pageIndex: state.flow.pageIndex,
        section: location.section, firstOnPage: band !== endnoteBand,
        bandTranslationPt: Object.freeze({ xPt: 0, yPt: state.flow.cursorBlockPt }),
        continuing: endnoteFlowStarted,
        container: { id: band, kind: 'endnote', capacity: 'unbounded', bounds: {
          xPt: location.availableBounds.xPt, yPt: 0,
          widthPt: location.availableBounds.widthPt, heightPt: location.section.geometry.pageHeight,
        } },
      }, cursor, location.availableBounds.heightPt, partitionFootnote);
      if (!partition) {
        if (state.flow.cursorBlockPt === state.flow.regionStartBlockPt && state.footnoteReservePt === 0) {
          throw new NoteCapacityExceededError('endnote', state.flow.pageIndex, band);
        }
        yield* commitTransition(advanceColumnOrPage(state.flow, 'overflow'), input.sequence.length);
        continue;
      }
      const placed = translateNoteLayout(partition.fragment, {
        xPt: 0, yPt: state.flow.cursorBlockPt - partition.fragment.flowBounds.yPt,
      });
      const retained = endnoteLayoutsByPage.get(state.flow.pageIndex) ?? [];
      retained.push(placed);
      endnoteLayoutsByPage.set(state.flow.pageIndex, retained);
      endnoteBand = band;
      endnoteFlowStarted = true;
      state = Object.freeze({ ...state, flow: placeFlowNode(state.flow, placed, placed.advancePt).state });
      cursor = partition.nextCursor;
      complete = cursor === null;
      if (!complete) yield* commitTransition(advanceColumnOrPage(state.flow, 'overflow'), input.sequence.length);
      observer?.onPageDrafts?.(state.pages);
      yield state.pages.length;
      if (observer?.shouldPublish(state.pages.length)) {
        observer.publish(paginationPassResult(state, owners, session, allocations,
          footnoteReserveByPage, footnoteLayoutsByPage, footnoteHostShiftsByPage,
          terminalDiagnostic, anchorInputs, serializedAnchorInputs, frameObservations,
          frameAnchorObservations, endnoteLayoutsByPage), input.sequence.length);
      }
    }
  }
  // A note may outlive the final body block; consume the generator so its
  // physical-page transition commits and the pending line cursor advances.
  while (pendingFootnoteContinuations.size > 0) {
    yield* commitTransition(
      // A column band continues in the next column of the page first
      // (WORD_PER_COLUMN_FOOTNOTE_BANDS).
      perColumnFootnoteBands()
        ? advanceColumnOrPage(state.flow, 'overflow')
        : advanceToPage(state.flow, state.flow.section, 'overflow'),
      input.sequence.length,
    );
    // Text-only tail transitions have no anchor prescan to suspend them.
    // Publish a safe page draft/checkpoint after each admitted note fragment,
    // so slicing and cancellation do not wait for the entire remaining note.
    observer?.onPageDrafts?.(state.pages);
    yield state.pages.length;
  }
  const reservedPages = new Set([
    ...footnoteReserveByPage.keys(),
    ...footnoteLayoutsByPage.keys(),
  ]);
  for (const pageIndex of reservedPages) {
    const reservePt = footnoteReserveByPage.get(pageIndex) ?? 0;
    const retained = footnoteLayoutsByPage.get(pageIndex) ?? [];
    const retainedAdvancePt = retained[0]?.pageArea?.heightPt
      ?? retained.reduce((sum, note) => sum + note.advancePt, 0);
    if (reservePt !== retainedAdvancePt) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `Page ${pageIndex} footnote reserve ${reservePt} does not equal retained advance ${retainedAdvancePt}`,
      );
    }
  }
  flushAnchorInputs();
  return paginationPassResult(
    state,
    owners,
    session,
    allocations,
    footnoteReserveByPage,
    footnoteLayoutsByPage,
    footnoteHostShiftsByPage,
    terminalDiagnostic,
    anchorInputs,
    serializedAnchorInputs,
    frameObservations,
    frameAnchorObservations,
    endnoteLayoutsByPage,
  );
}

type BandTranslation = Readonly<{ xPt: number; yPt: number }>;

/**
 * A header or footer story and the band translation that places it: the
 * header's flow top at headerDistance, the footer's flow end at footerDistance
 * above the page bottom. A band-dependent story (StoryLayout.bandDependent: a
 * page- or margin-anchored root cell-owner host, or positioned tables of its
 * tables) places that content on the page through the translation it is
 * given, and a footer's translation depends on its extent, so such a story is
 * laid out again until it is laid out with exactly the translation it
 * receives (bounded; a cycle or exhaustion fails closed). Any other story is
 * laid out once, unchanged.
 *
 * The translation moves along y only, so its fixed point is a root of one
 * scalar residual, which {@link solveExactTranslation} brackets and narrows,
 * still accepting only an exact pass. A residual that changes sign by a jump
 * alone (e.g. a host whose notBeside/none exclusion starts to displace a
 * following story paragraph exactly where the band would place it clear)
 * has no fixed point and fails closed.
 */
function layoutBandStory(
  layoutStory: (band?: BandTranslation) => StoryLayout,
  kind: 'header' | 'footer',
  geometry: Readonly<{ pageHeight: number; headerDistance: number; footerDistance: number }>,
): Readonly<{ story: StoryLayout; translation: BandTranslation }> {
  const translationOf = (story: StoryLayout): BandTranslation => Object.freeze({
    xPt: 0,
    yPt: (kind === 'header'
      ? geometry.headerDistance
      : geometry.pageHeight - geometry.footerDistance - story.advancePt) - story.flowBounds.yPt,
  });
  const unbanded = layoutStory();
  if (!unbanded.bandDependent) {
    return { story: unbanded, translation: translationOf(unbanded) };
  }
  try {
    // The limit is a resource guard only (evaluations); exact equality of the
    // translation a pass was laid out with and the one it implies accepts.
    return solveExactTranslation(translationOf(unbanded).yPt, (yPt) => {
      const story = layoutStory(Object.freeze({ xPt: 0, yPt }));
      const translation = translationOf(story);
      return { value: { story, translation }, implied: translation.yPt };
    }, 16);
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `${kind} band translation did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
}

/**
 * The composed position of a page's footnotes. ECMA-376 17.11.21 / 17.18.34:
 * pageBottom notes use the physical page's reserved body edge (earlier
 * continuous regions end inside the body and must not pull the page-wide
 * note band into preceding text); notes stack upward from it in document
 * order. `topsPt` is each note's page-final flow top.
 */
function footnoteStack(
  page: LayoutPage,
  notes: readonly NoteLayout[],
): Readonly<{ advancePt: number; targetTopPt: number; topsPt: readonly number[] }> {
  // ECMA-376 §17.11.21 defaults footnotes to the current page's bottom.
  const blockEndPt = page.sectionRegions.at(-1)?.blockEndPt
    ?? Math.max(
      0,
      page.section.geometry.pageHeight - Math.abs(page.section.geometry.marginBottom),
    );
  if (notes.some(note => note.pageArea !== undefined)) {
    const heightPt = notes[0]!.pageArea!.heightPt;
    if (!notes.every(note => note.pageArea?.heightPt === heightPt)) {
      throw new Error('A page-wide note area must own one common reserve');
    }
    const targetTopPt = blockEndPt - heightPt;
    return { advancePt: heightPt, targetTopPt,
      topsPt: Object.freeze(notes.map(note => targetTopPt + note.pageArea!.offsetPt)),
    };
  }
  // Each note band (its container, WORD_PER_COLUMN_FOOTNOTE_BANDS) stacks on
  // its own from the body bottom; a page band carries every column band.
  const bands = new Map<string, number[]>();
  notes.forEach((note, index) => {
    const members = bands.get(note.flowDomainId);
    if (members) members.push(index);
    else bands.set(note.flowDomainId, [index]);
  });
  const pageBand = [...bands.keys()].find((band) => !band.includes(':band:'));
  const pageStack = new NoteStack();
  (pageBand === undefined ? [] : bands.get(pageBand)!).forEach((index) => (
    pageStack.append(notes[index]!.advancePt)));
  const pageTopPt = blockEndPt - pageStack.advancePt;
  const topsPt = new Array<number>(notes.length);
  let targetTopPt = pageTopPt;
  for (const [band, members] of bands) {
    const isPageBand = band === pageBand;
    const stack = isPageBand ? pageStack : new NoteStack();
    if (!isPageBand) members.forEach((index) => stack.append(notes[index]!.advancePt));
    const bottomPt = isPageBand ? blockEndPt : pageTopPt;
    const topPt = bottomPt - stack.advancePt;
    targetTopPt = Math.min(targetTopPt, topPt);
    members.forEach((index, order) => { topsPt[index] = stack.topAt(topPt, order); });
  }
  const advancePt = bands.size === 1 && pageBand !== undefined
    ? pageStack.advancePt
    : blockEndPt - targetTopPt;
  return { advancePt, targetTopPt, topsPt };
}

/**
 * Footnote band plan (library policy): the page-final flow top of each
 * footnote whose story is band dependent (StoryLayout.bandDependent: its
 * tables hold §17.4.57 positioned tables, or text boxes whose stories hold
 * them; note roots elect no cell-owner carrier, table-owner-runs.ts
 * tableRowsElectCarriers). A footnote is
 * laid out when its reference is admitted, before later references on the
 * page fix the stack, so its page position is known only once the page is
 * composed. The reserve convergence carries this observation to the next
 * pass, which lays the note out with exactly that band, and accepts only an
 * exact fixed point (a pass composing every such note where it was laid out
 * for); a cycle or exhaustion fails closed there. Vertical pages keep the
 * plan too (composed tops are logical-frame positions), although they keep
 * no header/footer reserve.
 *
 * A continued tail is projected from its whole-note acquisition by its
 * source cut (FootnotePartition.hostShiftPt), so its content reaches the
 * page through the composed top and that shift together. The plan is that
 * sum, keyed by band and note together because a continuation can share a
 * physical page with its head: the acquisition's flow top, the band its next
 * acquisition is given. A box story in the tail then resolves its page
 * frames through the translation its anchor paragraph finally receives, so a
 * §17.4.57 page-anchored table in it lands on the authored page frame while
 * its text-anchored content follows the paragraph.
 */
function footnoteBandPlan(
  page: LayoutPage,
  notes: readonly NoteLayout[],
  hostShifts: ReadonlyMap<string, number> | undefined,
): Readonly<Record<string, number>> | undefined {
  if (!notes.some((note) => note.story.bandDependent)) return undefined;
  const { topsPt } = footnoteStack(page, notes);
  // ast-grep-ignore: no-plain-object-string-map -- keys are library-generated note band keys.
  const plan: Record<string, number> = {};
  notes.forEach((note, index) => {
    const key = noteBandKey(note.flowDomainId, note.source.storyInstance);
    if (note.story.bandDependent) plan[key] = topsPt[index]! + (hostShifts?.get(key) ?? 0);
  });
  return Object.freeze(plan);
}

/** ECMA-376 §17.6.11: page opening and final reserve observation use the same
 * selected story, band solve and signed-margin overflow rule. Unknown NUMPAGES
 * remains a pass input; on-demand measurement does not bypass exact convergence. */
function measurePageHeaderFooterReserve(
  session: BodyLayoutSession,
  context: PageStoryContext,
): HeaderFooterReserve {
  const { owner, pageIndex, section, firstPageOfSection, pageNumber } = context;
  // Vertical stories paint in physical page space; their overflow must not
  // charge the logical body interval. Use the resolved frame, including native
  // flow facts, rather than the text-direction token alone.
  if (sectionWritingMode(section) !== 'horizontal-tb') return { top: 0, bottom: 0 };
  const inlineExtentPt = Math.max(
    0,
    section.geometry.pageWidth
      - Math.abs(section.geometry.marginLeft)
      - Math.abs(section.geometry.marginRight),
  );
  const measure = (kind: 'header' | 'footer') => {
    const source = selectedHeaderFooterStory(
      kind === 'header' ? owner.headers : owner.footers,
      {
        titlePage: owner.titlePage,
        firstPageOfSection,
        evenAndOddHeaders: owner.evenAndOddHeaders,
        displayPageNumber: pageNumber.displayNumber,
      },
    );
    if (source === null) return 0;
    const layoutStory = session.layoutStory;
    if (!layoutStory) {
      throw new Error('Header/footer story layout requires a story-capable layout session');
    }
    // The reserve measures the story exactly as composition lays it out.
    const { story } = layoutBandStory((band) => layoutStory({
      source,
      pageStoryOwner: owner.sectionOccurrenceId,
      pageIndex,
      pageNumber: { displayNumber: pageNumber.displayNumber, format: pageNumber.format },
      section,
      container: {
        id: `story:${kind}:page:${pageIndex}`,
        kind,
        bounds: {
          xPt: Math.abs(section.geometry.marginLeft),
          yPt: 0,
          widthPt: inlineExtentPt,
          heightPt: section.geometry.pageHeight,
        },
      },
      ...(band ? { bandTranslationPt: band } : {}),
    }), kind, section.geometry);
    return kind === 'header' ? headerStoryBodyReserveExtentPt(story) : story.advancePt;
  };
  return Object.freeze({
    top: headerFooterOverflowReservePt(
      measure('header'),
      section.geometry.marginTop,
      section.geometry.headerDistance,
    ),
    bottom: headerFooterOverflowReservePt(
      measure('footer'),
      section.geometry.marginBottom,
      section.geometry.footerDistance,
    ),
  });
}

function headerFooterReserves(
  pass: BodyPaginationPassResult,
): readonly HeaderFooterReserve[] {
  return Object.freeze(pass.layout.pages.map((page, pageIndex) => {
    if (page.parityBlank) return Object.freeze({ top: 0, bottom: 0 });
    // A footnote's composed top is its page position in the section's logical
    // frame in every writing mode (composePageStories stacks it there).
    const footnoteTopsPt = footnoteBandPlan(
      page,
      pass.footnoteLayoutsByPage.get(page.pageIndex) ?? [],
      pass.footnoteHostShiftsByPage.get(page.pageIndex),
    );
    const context = pass.pageStoryContexts[pageIndex]!;
    return Object.freeze({
      ...measurePageHeaderFooterReserve(pass.session, context),
      ...(footnoteTopsPt ? { footnoteTopsPt } : {}),
      observedPage: { owner: context.owner.sectionOccurrenceId,
        firstPageOfSection: context.firstPageOfSection,
        parityBlank: false, pageNumber: context.pageNumber,
        ...(context.firstContentPageIndex === undefined ? {}
          : { firstContentPageIndex: context.firstContentPageIndex }) },
    });
  }));
}

function composePageStories(
  layout: DocumentLayout,
  session: BodyLayoutSession,
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
  footnotesByPage: ReadonlyMap<number, readonly NoteLayout[]>,
  endnotesByPage: ReadonlyMap<number, readonly NoteLayout[]>,
): DocumentLayout {
  const storyDiagnostics: LayoutDiagnostic[] = [];
  const visited = new WeakSet<object>();
  const pages = layout.pages.map((page, pageIndex) => {
    if (page.parityBlank) return page;
    const owner = owners.get(page.sectionOccurrenceId);
    if (!owner) throw new Error(`Unknown body section ${page.sectionOccurrenceId}`);
    if (!session.layoutStory) {
      const hasPageStories = Object.values(owner.headers).some((source) => source !== null)
        || Object.values(owner.footers).some((source) => source !== null)
        || (footnotesByPage.get(page.pageIndex)?.length ?? 0) > 0
        || (endnotesByPage.get(page.pageIndex)?.length ?? 0) > 0;
      if (!hasPageStories) return page;
      throw new Error('Page-story composition requires a story-capable layout session');
    }
    const vertical = sectionWritingMode(page.section) !== 'horizontal-tb';
    const geometry = vertical
      // `word-vertical-section-physical-header-footer`
      ? physicalSectionGeometry(page.section.geometry, page.section.sectionFlow)
      : page.section.geometry;
    const inlineStartPt = Math.abs(geometry.marginLeft);
    const inlineExtentPt = Math.max(
      0,
      geometry.pageWidth - Math.abs(geometry.marginLeft) - Math.abs(geometry.marginRight),
    );
    const coordinateSpace = vertical ? 'upright-physical' as const : 'section-logical' as const;
    // Upright physical page stories are horizontal, so a native section frame
    // does not carry into them.
    const { sectionFlow: _sectionFlow, ...pageSection } = page.section;
    const pageStorySection: DeepReadonly<SectionLayoutContext> = vertical
      ? Object.freeze({
          ...pageSection,
          geometry: Object.freeze({ ...geometry }),
          columns: Object.freeze([Object.freeze({
            xPt: inlineStartPt,
            wPt: inlineExtentPt,
          })]),
          textDirection: 'lrTb',
        })
      : page.section;
    const sourceFor = (kind: 'header' | 'footer') => selectedHeaderFooterStory(
      kind === 'header' ? owner.headers : owner.footers,
      {
        titlePage: owner.titlePage,
        firstPageOfSection: isFirstSectionOwnedPage(layout.pages, pageIndex),
        evenAndOddHeaders: owner.evenAndOddHeaders,
        displayPageNumber: page.pageNumber.displayNumber,
      },
    );
    const acquire = (kind: 'header' | 'footer') => {
      const source = sourceFor(kind);
      if (source === null) return null;
      const { story, translation } = layoutBandStory((band) => session.layoutStory!({
        source,
        pageStoryOwner: owner.sectionOccurrenceId,
        pageIndex: page.pageIndex,
        pageNumber: { displayNumber: page.pageNumber.displayNumber, format: page.pageNumber.format },
        section: pageStorySection,
        container: {
          id: `story:${kind}:page:${page.pageIndex}`,
          kind,
          bounds: {
            xPt: inlineStartPt,
            yPt: 0,
            widthPt: inlineExtentPt,
            heightPt: geometry.pageHeight,
          },
        },
        ...(band ? { bandTranslationPt: band } : {}),
      }), kind, geometry);
      return translateStoryLayout(story, translation);
    };
    const header = acquire('header');
    const footer = acquire('footer');
    const retainedNotes = footnotesByPage.get(page.pageIndex) ?? [];
    const noteRegion = page.sectionRegions.at(-1);
    const noteStack = footnoteStack(page, retainedNotes);
    const noteAdvancePt = noteStack.advancePt;
    const noteTargetTopPt = noteStack.targetTopPt;
    const notes = [
      ...(endnotesByPage.get(page.pageIndex) ?? []),
      ...retainedNotes.map((note, index) => translateNoteLayout(note, {
        xPt: 0,
        yPt: noteStack.topsPt[index]! - note.flowBounds.yPt,
      })),
    ];
    const noteInlineStartPt = notes.length === 0
      ? 0
      : Math.min(...notes.map((note) => note.flowBounds.xPt));
    const noteInlineEndPt = notes.length === 0
      ? 0
      : Math.max(...notes.map((note) => note.flowBounds.xPt + note.flowBounds.widthPt));
    const noteLogicalBounds = Object.freeze({
      xPt: noteInlineStartPt,
      yPt: noteTargetTopPt,
      widthPt: noteInlineEndPt - noteInlineStartPt,
      heightPt: noteAdvancePt,
    });
    const notePhysicalBounds = noteRegion
      ? Object.freeze(transformRect(
          noteRegion.coordinateSpace.logicalToPhysical,
          noteLogicalBounds,
        ))
      : noteLogicalBounds;
    const storyDomains = [
      ...(header ? [Object.freeze({
        id: `story:header:page:${page.pageIndex}`,
        kind: 'header' as const,
        logicalBounds: Object.freeze({
          xPt: inlineStartPt,
          yPt: header.flowBounds.yPt,
          widthPt: inlineExtentPt,
          heightPt: header.advancePt,
        }),
        physicalBounds: Object.freeze({
          xPt: inlineStartPt,
          yPt: header.flowBounds.yPt,
          widthPt: inlineExtentPt,
          heightPt: header.advancePt,
        }),
      })] : []),
      ...(notes.length > 0 && notes.every((note) => note.flowDomainId === `notes:page:${page.pageIndex}`)
        ? [Object.freeze({
          id: `notes:page:${page.pageIndex}`,
          kind: 'footnote' as const,
          ...(noteRegion ? { sectionRegionId: noteRegion.id } : {}),
          logicalBounds: noteLogicalBounds,
          physicalBounds: notePhysicalBounds,
        })]
        // WORD_PER_COLUMN_FOOTNOTE_BANDS: one story domain per note band.
        : [...new Set(notes.map((note) => note.flowDomainId))].map((bandId) => {
          const members = notes.filter((note) => note.flowDomainId === bandId);
          const left = Math.min(...members.map((note) => note.flowBounds.xPt));
          const top = Math.min(...members.map((note) => note.flowBounds.yPt));
          const logicalBounds = Object.freeze({
            xPt: left,
            yPt: top,
            widthPt: Math.max(...members.map((note) => note.flowBounds.xPt + note.flowBounds.widthPt)) - left,
            heightPt: Math.max(...members.map((note) => note.flowBounds.yPt + note.flowBounds.heightPt)) - top,
          });
          return Object.freeze({
            id: bandId,
            kind: members[0]!.source.story === 'endnote' ? 'endnote' as const : 'footnote' as const,
            ...(noteRegion ? { sectionRegionId: noteRegion.id } : {}),
            logicalBounds,
            physicalBounds: noteRegion
              ? Object.freeze(transformRect(noteRegion.coordinateSpace.logicalToPhysical, logicalBounds))
              : logicalBounds,
          });
        })),
      ...(footer ? [Object.freeze({
        id: `story:footer:page:${page.pageIndex}`,
        kind: 'footer' as const,
        logicalBounds: Object.freeze({
          xPt: inlineStartPt,
          yPt: footer.flowBounds.yPt,
          widthPt: inlineExtentPt,
          heightPt: footer.advancePt,
        }),
        physicalBounds: Object.freeze({
          xPt: inlineStartPt,
          yPt: footer.flowBounds.yPt,
          widthPt: inlineExtentPt,
          heightPt: footer.advancePt,
        }),
      })] : []),
    ];
    const existing = page.layers.roots.map((entry): PageLayerNode => entry);
    const firstNonLeading = existing.findIndex((entry) =>
      entry.layer !== 'background' && entry.layer !== 'behindText');
    const headerIndex = firstNonLeading < 0 ? existing.length : firstNonLeading;
    const withHeader = [
      ...existing.slice(0, headerIndex),
      ...(header?.blocks.map((node): PageLayerNode => ({
        layer: 'header', node, coordinateSpace,
      })) ?? []),
      ...existing.slice(headerIndex),
    ];
    let lastBodyIndex = -1;
    for (let index = 0; index < withHeader.length; index += 1) {
      if (withHeader[index]!.layer === 'body') lastBodyIndex = index;
    }
    const noteIndex = lastBodyIndex < 0 ? withHeader.length : lastBodyIndex + 1;
    const entries: PageLayerNode[] = [
      ...withHeader.slice(0, noteIndex),
      ...notes.map((node): PageLayerNode => ({
        layer: 'notes', node, coordinateSpace: 'section-logical',
      })),
      ...withHeader.slice(noteIndex),
      ...(footer?.blocks.map((node): PageLayerNode => ({
        layer: 'footer', node, coordinateSpace,
      })) ?? []),
    ];
    // The body pass collected only body diagnostics; these stories are
    // composed here, so their own diagnostics are collected here.
    storyDiagnostics.push(...[...(header?.blocks ?? []), ...notes, ...(footer?.blocks ?? [])]
      .flatMap((node) => nestedStoryDiagnostics(node, visited)));
    return Object.freeze({
      ...page,
      flowDomains: Object.freeze([...page.flowDomains, ...storyDomains]),
      layers: createPageLayers(entries),
      readingOrder: Object.freeze([
        ...(header?.blocks.map((node) => node.id) ?? []),
        ...page.readingOrder,
        ...notes.map((note) => note.id),
        ...(footer?.blocks.map((node) => node.id) ?? []),
      ]),
    });
  });
  return Object.freeze({
    ...layout,
    pages: Object.freeze(pages),
    ...(storyDiagnostics.length > 0
      ? { diagnostics: Object.freeze([...layout.diagnostics, ...storyDiagnostics]) }
      : {}),
  });
}

function appendUnsupportedNotePositionDiagnostic(
  layout: DocumentLayout,
  kind: 'footnote' | 'endnote',
  position: string,
  fallback: 'pageBottom' | 'docEnd',
): DocumentLayout {
  return Object.freeze({
    ...layout,
    diagnostics: Object.freeze([...layout.diagnostics, Object.freeze({
      code: 'UNSUPPORTED_FEATURE' as const,
      severity: 'error' as const,
      message: `Unsupported ${kind} position ${JSON.stringify(position)}; `
        + `retained layout uses the ${fallback} fallback`,
    })]),
  });
}

/** The flow domain in which the section region holding `flowDomainId` opens
 * on `page`; its page-start prescan is where page-owned anchors register. */
function regionFlowDomainOpening(page: LayoutPage, flowDomainId: string): string {
  return page.sectionRegions.find((region) => region.flowDomainIds.includes(flowDomainId))
    ?.flowDomainIds[0] ?? flowDomainId;
}

/** Flow order and anchor line of each page-owned drawing in a pass. */
function pageOwnedDrawingFlowPositions(layout: DocumentLayout) {
  const positions = new Map<string, Readonly<{ order: number; lineKey: string }>>();
  let order = 0;
  for (const page of layout.pages) {
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph' || node.drawings.length === 0) continue;
      pageOwnedAnchorKeysByLine(node).forEach((keys, lineIndex) => {
        for (const key of keys) {
          positions.set(key, Object.freeze({ order, lineKey: `${node.id}#${lineIndex}` }));
          order += 1;
        }
      });
    }
  }
  return positions;
}

function pageAnchorDestinationPlan(layout: DocumentLayout) {
  const destinations = new Map<string, PageWrapDestination>();
  for (const page of layout.pages) {
    for (const node of page.layers.body) {
      if (node.kind === 'table' && !node.ordinaryFlow
        && node.sectionFlowOwnership === 'page' && node.cellOwnerHost !== true) {
        // §17.4.57 permits a page-positioned table to exclude text that
        // precedes it in source order. The first pass owns its actual page and
        // fragment extent; the next pass reserves exactly that page-local box.
        // Continuations switch to text-owned flow, so only the first root
        // fragment has page ownership and cursor (row 0, fragment 0).
        // Scope: this resolves an accepted page/margin root's collision with
        // preceding visible lines. It does not yet reinterpret §17.4.57's
        // logical anchor at the following regular paragraph when that owner
        // differs from the table's current source-page assignment.
        const occurrenceId = bodyRootFloatingTablePlacementKey(node.source, page.pageIndex, 0, 0);
        destinations.set(`table:${sourceKey(node.source)}`,
          Object.freeze({
            kind: 'floating-table',
            occurrenceId,
            tableSource: node.source,
            bounds: Object.freeze({ ...node.flowBounds }),
            pageIndex: page.pageIndex,
            flowDomainId: node.flowDomainId,
          }));
        continue;
      }
      if (node.kind !== 'paragraph') continue;
      for (const drawing of node.drawings) {
        const anchor = drawing.anchorLayer;
        if (!anchor
          || anchor.horizontalOwnership !== 'page'
          || anchor.verticalOwnership !== 'page') continue;
        const occurrenceId = anchor.acquisitionOccurrenceId ?? anchor.occurrenceId;
        destinations.set(occurrenceId, Object.freeze({
          kind: 'drawing',
          occurrenceId,
          paragraphSource: node.source,
          pageIndex: page.pageIndex,
          // word-page-anchor-region-registration: register the drawing where
          // its section region opens on the page, so earlier columns of the
          // region wrap around it too.
          flowDomainId: regionFlowDomainOpening(page, node.flowDomainId),
        }));
      }
    }
  }
  return destinations;
}

/** The anchors a pass applying `plan` prescans when it opens this flow domain. */
function plannedPageStartAnchors(
  plan: ReadonlyMap<string, PageWrapDestination>,
  pageIndex: number,
  flowDomainId: string,
): PageStartAnchors {
  return Object.freeze([...plan.values()]
    .filter((destination): destination is Exclude<PageWrapDestination, { kind: 'frame-deferral' }> => (
      destination.kind !== 'frame-deferral'
      && destination.pageIndex === pageIndex && destination.flowDomainId === flowDomainId
    ))
    .map((destination) => destination.kind === 'drawing'
      ? Object.freeze({
          kind: 'drawing' as const,
          occurrenceId: destination.occurrenceId,
          paragraphSource: destination.paragraphSource,
        })
      : destination.kind === 'host-drawing'
        ? Object.freeze({
            kind: 'host-drawing' as const,
            occurrenceId: destination.occurrenceId,
            paragraphSource: destination.paragraphSource,
            carry: destination.carry,
          })
        : destination.kind === 'frame'
        ? Object.freeze({
            kind: 'frame' as const,
            occurrenceId: destination.occurrenceId,
            paragraphSource: destination.paragraphSource,
            carry: destination.carry,
          })
        : Object.freeze({
            kind: 'floating-table' as const,
            occurrenceId: destination.occurrenceId,
            tableSource: destination.tableSource,
            bounds: destination.bounds,
          })));
}

function pageStartAnchorsIdentity(anchors: PageStartAnchors): string {
  return anchors.map((anchor) => (anchor.kind === 'drawing'
    ? `drawing|${anchor.occurrenceId}|${sourceKey(anchor.paragraphSource)}`
    : anchor.kind === 'host-drawing'
      ? `host|${anchor.occurrenceId}|${sourceKey(anchor.paragraphSource)}|${JSON.stringify(anchor.carry)}`
      : anchor.kind === 'frame'
      ? `frame|${anchor.occurrenceId}|${sourceKey(anchor.paragraphSource)}|${JSON.stringify(anchor.carry)}`
      : `table|${anchor.occurrenceId}|${sourceKey(anchor.tableSource)}|${anchor.bounds.xPt}|${
        anchor.bounds.yPt}|${anchor.bounds.widthPt}|${anchor.bounds.heightPt}`)).join('\n');
}

/**
 * WORD_LATER_ANCHOR_EARLIER_LINE_WRAP / WORD_MODE14_TIGHT_ANCHOR_TOP_TOUCH:
 * the paragraph-relative wrapping drawings of a pass that must also wrap
 * earlier lines of their page. A drawing qualifies when its exclusion starts
 * above its anchor paragraph, or, for a compatibility-mode-14 wrapTight
 * drawing, when it starts exactly at the paragraph top and the paragraph's
 * first line could not be placed there. Each carries the geometry resolved in
 * this pass (its first placement).
 */
function hostAnchorCarryPlan(layout: DocumentLayout): Map<string, PageWrapDestination> {
  const plan = new Map<string, PageWrapDestination>();
  for (const page of layout.pages) {
    // Exclusions of drawings anchored in earlier paragraphs of this page, with
    // the top of their anchor paragraph.
    const earlier: Array<Readonly<{ paragraphTopPt: number; bounds: LayoutRect }>> = [];
    // Line bands of the paragraphs already laid out on this page: a carry
    // that meets none of them cannot change the page.
    const earlierBands: Array<Readonly<{ topPt: number; bottomPt: number }>> = [];
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph') continue;
      const bands = node.lines.length > 0
        ? node.lines.map((line) => Object.freeze({
          topPt: line.bounds.yPt,
          bottomPt: line.bounds.yPt + line.bounds.heightPt,
        }))
        : [Object.freeze({
          topPt: node.flowBounds.yPt,
          bottomPt: node.flowBounds.yPt + node.flowBounds.heightPt,
        })];
      if (node.drawings.length === 0) {
        earlierBands.push(...bands);
        continue;
      }
      const owned = new Set((node.anchorCollisions ?? []).map((entry) => entry.occurrenceId));
      const ownedExclusions = node.exclusions.filter((exclusion) => (
        exclusion.anchorOccurrenceId !== undefined && owned.has(exclusion.anchorOccurrenceId)
      ));
      const pending = () => ownedExclusions.forEach((exclusion) => earlier.push(Object.freeze({
        paragraphTopPt: node.flowBounds.yPt,
        bounds: exclusion.bounds,
      })));
      if (node.continuation?.continuesFromPrevious) {
        pending();
        earlierBands.push(...bands);
        continue;
      }
      const exclusions = new Map(node.exclusions.flatMap((exclusion) => (
        exclusion.anchorOccurrenceId ? [[exclusion.anchorOccurrenceId, exclusion] as const] : []
      )));
      // Retained ids are projected; acquisition and the page registry use the
      // acquisition occurrence id.
      const acquisitionIds = new Map(node.drawings.flatMap((drawing) => (
        drawing.anchorLayer
          ? [[drawing.anchorLayer.occurrenceId,
            drawing.anchorLayer.acquisitionOccurrenceId ?? drawing.anchorLayer.occurrenceId] as const]
          : []
      )));
      const frames = new Map((node.anchorFrames ?? []).flatMap((frame) => (
        frame.status === 'resolved' ? [[frame.occurrenceId, frame] as const] : []
      )));
      const paragraphTopPt = node.flowBounds.yPt;
      const firstLine = node.lines[0];
      const firstLinePushed = firstLine !== undefined
        && firstLine.bounds.yPt > paragraphTopPt + node.spacing.beforePt;
      for (const collision of node.anchorCollisions ?? []) {
        if (collision.verticalOwnership !== 'host') continue;
        const frame = frames.get(collision.occurrenceId);
        const exclusion = exclusions.get(collision.occurrenceId);
        if (!frame || !exclusion || frame.geometry.wrap.kind === 'none') continue;
        const topPt = exclusion.bounds.yPt;
        const contentTopPt = paragraphTopPt + node.spacing.beforePt;
        // The touch reaches back to the anchor paragraph of the earliest other
        // drawing that also meets the first-line band where the line started.
        const touchFromPt = exclusion.wordMode14TightAnchor === true
          && firstLinePushed
          && topPt === contentTopPt
          && firstLine !== undefined
          ? earlier.find((entry) => (
            entry.bounds.yPt < contentTopPt + firstLine.bounds.heightPt
              && entry.bounds.yPt + entry.bounds.heightPt > contentTopPt
          ))?.paragraphTopPt
          : undefined;
        const touch = touchFromPt !== undefined;
        if (!(topPt < paragraphTopPt) && !touch) continue;
        const carry: CarriedHostAnchor = Object.freeze({
          bounds: Object.freeze({ ...collision.bounds }),
          exclusionBounds: Object.freeze({ ...exclusion.bounds }),
          horizontalOwnership: collision.horizontalOwnership,
          verticalOwnership: collision.verticalOwnership,
          wrap: exclusion.wrap,
          wrapSide: frame.geometry.wrap.side,
          wrapDistances: Object.freeze({ ...frame.geometry.wrap.distances }),
          ...(frame.geometry.wrap.polygon
            ? { wrapPolygon: Object.freeze([...frame.geometry.wrap.polygon.points]) }
            : {}),
          ...(touch ? { topEdgeInclusiveFromYPt: touchFromPt } : {}),
          ...(exclusion.anchorLineExemptTopPt === undefined
            ? {} : { anchorLineExemptTopPt: exclusion.anchorLineExemptTopPt }),
        });
        // The line check's own predicate, so a carry is kept exactly when it
        // can change an earlier line.
        const meets = createFloatLineBandPredicate(carriedFloatRect(carry));
        if (!earlierBands.some((band) => meets(band.topPt, band.bottomPt - band.topPt))) continue;
        const occurrenceId = acquisitionIds.get(collision.occurrenceId);
        if (occurrenceId === undefined) continue;
        plan.set(occurrenceId, Object.freeze({
          kind: 'host-drawing',
          occurrenceId,
          paragraphSource: node.source,
          pageIndex: page.pageIndex,
          flowDomainId: node.flowDomainId,
          carry,
        }));
      }
      pending();
      earlierBands.push(...bands);
    }
  }
  return plan;
}

/** The wrap geometry the page registry derives from a carried drawing (see
 * commitBodyFlowRegistryDelta and paragraphWrapExclusions). */
function carriedFloatRect(carry: CarriedHostAnchor): FloatRect {
  const bounds = carry.exclusionBounds;
  return {
    kind: 'shape',
    mode: carry.wrap === 'topAndBottom' ? 'topAndBottom' : 'square',
    authoredWrap: carry.wrap,
    ...(carry.wrapPolygon ? { wrapPolygon: carry.wrapPolygon } : {}),
    imageKey: 'carried-anchor',
    imageX: carry.bounds.xPt,
    imageY: carry.bounds.yPt,
    imageW: carry.bounds.widthPt,
    imageH: carry.bounds.heightPt,
    xLeft: bounds.xPt,
    xRight: bounds.xPt + bounds.widthPt,
    yTop: bounds.yPt,
    yBottom: bounds.yPt + bounds.heightPt,
    side: carry.wrapSide ?? 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
    paraId: 0,
    ...(carry.topEdgeInclusiveFromYPt === undefined
      ? {} : { topEdgeInclusiveFromYPt: carry.topEdgeInclusiveFromYPt }),
    ...(carry.anchorLineExemptTopPt === undefined
      ? {} : { exemptLineTopPt: carry.anchorLineExemptTopPt }),
  };
}

/** The carried drawings of `applied` that stayed on their page with their
 * carried frame: a carried drawing whose anchor paragraph lands elsewhere, or
 * that §20.4.2.3 collision avoidance moved off its carried frame, is dropped
 * for good. */
function keptHostAnchorCarries(
  applied: ReadonlyMap<string, PageWrapDestination> | null,
  layout: DocumentLayout,
): Map<string, PageWrapDestination> {
  const kept = new Map<string, PageWrapDestination>();
  if (!applied) return kept;
  const landed = new Map<string, Readonly<{ at: string; bounds: LayoutRect | undefined }>>();
  for (const page of layout.pages) {
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph') continue;
      const collisions = new Map((node.anchorCollisions ?? []).map((entry) => [entry.occurrenceId, entry.bounds] as const));
      for (const drawing of node.drawings) {
        if (!drawing.anchorLayer) continue;
        landed.set(
          drawing.anchorLayer.acquisitionOccurrenceId ?? drawing.anchorLayer.occurrenceId,
          Object.freeze({
            at: `${page.pageIndex}|${node.flowDomainId}`,
            bounds: collisions.get(drawing.anchorLayer.occurrenceId),
          }),
        );
      }
    }
  }
  const sameRect = (left: LayoutRect | undefined, right: LayoutRect) => left !== undefined
    && left.xPt === right.xPt && left.yPt === right.yPt
    && left.widthPt === right.widthPt && left.heightPt === right.heightPt;
  for (const [key, destination] of applied) {
    if (destination.kind !== 'host-drawing') continue;
    const observed = landed.get(key);
    if (observed?.at === `${destination.pageIndex}|${destination.flowDomainId}`
      && sameRect(observed.bounds, destination.carry.bounds)) {
      kept.set(key, destination);
    }
  }
  return kept;
}

/** Lowest page index holding a carry candidate, bounding what the carry-free
 * pass may publish. */
function firstHostCarryPage(layout: DocumentLayout): number {
  let first = Number.POSITIVE_INFINITY;
  for (const destination of hostAnchorCarryPlan(layout).values()) {
    first = Math.min(first, destination.pageIndex);
  }
  return first;
}

/**
 * The page-index bound below which a snapshot of an in-progress page-anchor
 * pass is final: no later pass of this convergence run can change those pages.
 *
 * A pass is a deterministic function of the body input, the reserves and the
 * balance plan — fixed for the whole run — and of the values carried between
 * passes: the anchor plan, the proven minimum table pages and the proven
 * anchor-line deferrals. The pass reads the carry only at the events it
 * reports through `onPageAnchorInput`; a deferral check that does not apply
 * changes nothing and is not a read. Two passes whose event reads agree up to
 * some moment are therefore in the same state at that moment, including every
 * page closed by then.
 *
 * Pass k+1 applies the plan pass k observed (`pageAnchorDestinationPlan`),
 * except that `resolveAnchorLineTests` may retest drawings on a page N, and
 * may prove new deferrals for page N, only where a drawing registered on N
 * did not land on N. Pass k's prescan of N then already disagrees with the
 * observed plan (the drawing is observed later, or not yet placed), so both
 * only affect pages at or after a disagreement. Each read can be predicted
 * from closed pages of pass k:
 *
 * - A prescan of page p reads the plan's destinations on (p, flow domain). Once
 *   p is closed, pass k's observed destinations there are final, and the read
 *   agrees iff they equal what pass k prescanned.
 * - A page-owned table floor read at page r changes the flow only when the
 *   floor exceeds r. The next floor is absent, carried, or proven from a table
 *   whose destination changed. The read agrees in every later pass iff the
 *   current floor does not exceed r, the table's destination is unchanged,
 *   and that destination lies on a page that itself stays final. A table
 *   reached before the bound but placed at or after it could still move, and
 *   its next floor could then act at r, so it lowers the bound to r.
 * - An applied anchor-line deferral at page r ends r above the anchor line;
 *   the bound stops at r. A carried proof for page p is verified only against
 *   this pass's own reads up to its anchor line on p, which agree in the next
 *   pass by induction; a new proof for the same anchor and page needs a
 *   disagreeing prescan of p.
 *
 * By induction every later pass, including the converged one, reproduces
 * pages below the returned bound exactly. The bound never exceeds the live
 * page, whose content is still open. The first pass, which prescans source
 * order rather than a plan, is covered by the same comparison.
 *
 * A second pass that runs while the first is still discovering
 * (`LeadingSecondPass`) applies a plan that so far holds only the pages the
 * first pass has closed. A table missing from that partial plan may still be
 * placed by the first pass, so its prior destination is unknown rather than
 * absent, and the bound stops at its reach page (`appliedPlanComplete`).
 *
 * Scope: this proves stability across anchor convergence only. Later
 * continuous-section balancing and header/footer reserve or pagination-field
 * convergence can still replace a provisional publication (`exact:false`).
 */
function anchorStablePageLimit(
  events: readonly PageAnchorInputEvent[],
  appliedPlan: ReadonlyMap<string, PageWrapDestination> | null,
  snapshot: BodyPaginationPassResult,
  appliedPlanComplete = true,
): Readonly<{ limit: number; prescanDisagreementPage: number }> {
  const observed = pageAnchorDestinationPlan(snapshot.layout);
  let limit = livePageIndex(snapshot);
  // A disagreeing prescan on a closed page stays disagreeing for the rest of
  // this pass, so its page also caps every later snapshot of the pass.
  let prescanDisagreementPage = Number.POSITIVE_INFINITY;
  const reachedTables: Array<Readonly<{ reachedPage: number; placedPage: number }>> = [];
  for (const event of events) {
    // Event pages never decrease: the flow only advances.
    if (event.pageIndex >= limit) break;
    if (event.kind === 'prescan') {
      const next = plannedPageStartAnchors(observed, event.pageIndex, event.flowDomainId);
      if (pageStartAnchorsIdentity(next) !== pageStartAnchorsIdentity(event.anchors)) {
        limit = event.pageIndex;
        prescanDisagreementPage = event.pageIndex;
        break;
      }
      continue;
    }
    if (event.kind === 'anchor-line-deferral') {
      // A deferral ends its page above the anchor line. The pages before it
      // are covered by the reads above; this page and later ones are not.
      limit = event.pageIndex;
      break;
    }
    const placed = observed.get(event.key);
    const prior = appliedPlan?.get(event.key);
    if ((appliedPlan !== null && prior === undefined && !appliedPlanComplete)
      || (event.floor !== undefined && event.floor > event.pageIndex)
      || placed?.kind !== 'floating-table'
      || (prior !== undefined && JSON.stringify(prior) !== JSON.stringify(placed))) {
      limit = event.pageIndex;
      break;
    }
    reachedTables.push(Object.freeze({
      reachedPage: event.pageIndex,
      placedPage: placed.pageIndex,
    }));
  }
  for (let lowered = true; lowered;) {
    lowered = false;
    for (const table of reachedTables) {
      if (table.reachedPage < limit && table.placedPage >= limit) {
        limit = table.reachedPage;
        lowered = true;
      }
    }
  }
  return Object.freeze({ limit, prescanDisagreementPage });
}

function anchorCarryIdentity(
  plan: ReadonlyMap<string, unknown>,
  deferrals: AnchorLineDeferrals,
): string {
  return `${anchorPlanIdentity(plan)}\u0004${anchorLineDeferralsIdentity(deferrals)}`;
}

/**
 * The next pass's drawing plan and anchor-line deferrals after a planned pass
 * (see anchor-line-deferral.ts for the Word rule).
 *
 * On each page N, the first anchor line in flow order whose page-owned
 * anchors were registered on N but landed after N is Word's next test on N.
 * When N registered exactly those anchors plus the anchors confirmed on N, the
 * pass is Word's counterfactual and proves the deferral. Otherwise the next
 * pass retests them on N with exactly that registration. A line reached later
 * on N than a failed test is never tested on N, because the failed test ends
 * the page above it.
 */
function resolveAnchorLineTests(
  pass: BodyPaginationPassResult,
  observed: Map<string, PageWrapDestination>,
  previous: AnchorLineDeferrals,
): Readonly<{ plan: ReadonlyMap<string, PageWrapDestination>; deferrals: AnchorLineDeferrals }> {
  const events = pass.anchorInputs;
  const prescansByPage = new Map<number, number[]>();
  events.forEach((event, index) => {
    if (event.kind !== 'prescan') return;
    const indexes = prescansByPage.get(event.pageIndex);
    if (indexes) indexes.push(index);
    else prescansByPage.set(event.pageIndex, [index]);
  });
  let positions: ReturnType<typeof pageOwnedDrawingFlowPositions> | null = null;
  let plan: Map<string, PageWrapDestination> | null = null;
  let deferrals: Map<string, AnchorLineDeferralProof> | null = null;
  for (const [pageIndex, indexes] of prescansByPage) {
    const registeredAt = new Map<string, number>();
    for (const index of indexes) {
      const event = events[index] as Extract<PageAnchorInputEvent, { kind: 'prescan' }>;
      for (const anchor of event.anchors) {
        if (anchor.kind === 'drawing' && !registeredAt.has(anchor.occurrenceId)) {
          registeredAt.set(anchor.occurrenceId, index);
        }
      }
    }
    const pushed = [...registeredAt.keys()].filter((key) => (
      (observed.get(key)?.pageIndex ?? -1) > pageIndex
    ));
    if (pushed.length === 0) continue;
    positions ??= pageOwnedDrawingFlowPositions(pass.layout);
    let first: string | null = null;
    for (const key of pushed) {
      const order = positions.get(key)?.order;
      if (order === undefined) continue;
      if (first === null || order < positions.get(first)!.order) first = key;
    }
    if (first === null) continue;
    const lineKey = positions.get(first)!.lineKey;
    const line = [...positions].filter(([, position]) => position.lineKey === lineKey)
      .map(([key]) => key);
    const registrationIndex = registeredAt.get(first)!;
    // Anchors on one line are tested together, in one registration.
    if (line.some((key) => registeredAt.get(key) !== registrationIndex)) continue;
    const tested = new Set(line);
    const exact = indexes.every((index) => {
      const event = events[index] as Extract<PageAnchorInputEvent, { kind: 'prescan' }>;
      return pageStartAnchorsIdentity(event.anchors.filter((anchor) => !tested.has(anchor.occurrenceId)))
        === pageStartAnchorsIdentity(plannedPageStartAnchors(observed, pageIndex, event.flowDomainId));
    });
    if (exact) {
      const proof = createAnchorLineDeferralProof(
        line,
        pageIndex,
        events,
        pass.serializedAnchorInputs,
        registrationIndex,
        pageStartAnchorsIdentity,
      );
      deferrals ??= new Map(previous);
      line.forEach((key) => deferrals!.set(anchorLineDeferralKey(key, pageIndex), proof));
      continue;
    }
    plan ??= new Map(observed);
    const registration = events[registrationIndex] as Extract<PageAnchorInputEvent, { kind: 'prescan' }>;
    for (const key of line) {
      const destination = observed.get(key)!;
      plan.set(key, Object.freeze({
        ...destination,
        pageIndex,
        flowDomainId: registration.flowDomainId,
      }));
    }
  }
  return Object.freeze({ plan: plan ?? observed, deferrals: deferrals ?? previous });
}

function anchorPlanIdentity(plan: ReadonlyMap<string, unknown>): string {
  return JSON.stringify([...plan].sort(([left], [right]) => left.localeCompare(right)));
}

/** Source order of two page-owned table keys (`table:<story>:<instance>:<path>`):
 * body paths compare element by element, numerically. */
function compareTableSourceOrder(left: string, right: string): number {
  const path = (key: string) => key.slice(key.lastIndexOf(':') + 1).split('.').map(Number);
  const [a, b] = [path(left), path(right)];
  for (let index = 0; index < Math.min(a.length, b.length); index += 1) {
    if (a[index] !== b[index]) return a[index]! - b[index]!;
  }
  return a.length - b.length || (left < right ? -1 : left > right ? 1 : 0);
}

function changedAnchorKeys(
  applied: ReadonlyMap<string, PageWrapDestination>,
  observed: ReadonlyMap<string, PageWrapDestination>,
): ReadonlySet<string> {
  const changed = new Set<string>();
  for (const [key, destination] of applied) {
    if (JSON.stringify(destination) !== JSON.stringify(observed.get(key))) changed.add(key);
  }
  for (const key of observed.keys()) {
    if (!applied.has(key)) changed.add(key);
  }
  return changed;
}

/** What a header/footer reserve repagination reads of the previous pass. */
type ReserveRepaginationCarry = Readonly<{
  fieldContexts: ReturnType<typeof paginationFieldPageContexts>;
  anchorPlan: ReturnType<typeof pageAnchorDestinationPlan>;
}>;

/** An unseeded run's second anchor pass, already advanced while the first
 * pass ran (see `paginateOpeningPassWithLeadingSecond`). */
type LeadingSecondPass = Readonly<{
  /** Whether the carry the solver hands its second step is exactly the one
   * this pass has been applying. */
  appliesCarry(
    plan: ReadonlyMap<string, PageWrapDestination>,
    minimumTablePageBySource: ReadonlyMap<string, number>,
    deferrals: AnchorLineDeferrals,
  ): boolean;
  /** Completes the pass from where it stands. */
  finish(): PaginationSteps<BodyPaginationPassResult>;
}>;

const ANCHOR_PASS_BASE_LIMIT = 16;

/** Page numbering is not an input of `pageAnchorDestinationPlan`; a closed page
 * is finalized with this placeholder only to read its destinations. */
const CLOSED_PAGE_NUMBER = Object.freeze({
  displayNumber: 0,
  format: 'decimal',
  sectionOccurrenceId: '',
});

function* paginateBodyWithAnchorConvergenceSteps(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
  balancePlan: BodyBalancePlan,
  publisher?: BodyPagePublisher,
  seedPlan?: ReturnType<typeof pageAnchorDestinationPlan>,
): PaginationSteps<BodyPaginationPassResult> {
  let pageOwnedAnchorCount = 0;
  for (const entry of input.sequence) {
    if (entry.kind !== 'body-block') continue;
    pageOwnedAnchorCount += entry.block.kind === 'paragraph'
      ? entry.block.pageOwnedAnchorOccurrenceIds?.length ?? 0
      : entry.block.pageOwnedFloatingTable === true ? 1 : 0;
  }
  // Operational resource guard, not a claim about the state space. Pages
  // before the first disagreeing read are final (`anchorStablePageLimit`), and
  // once they are, one anchor line's Word test needs at most four passes:
  // observe the line, retest it with exactly the anchors confirmed before it,
  // prove the deferral, apply it. Anchor tests on later pages wait for the
  // pages before them, so the budget grows with the anchors, not a constant.
  let anchorPassLimit = ANCHOR_PASS_BASE_LIMIT + 4 * pageOwnedAnchorCount;
  // A carry-free pass may change from the first page that holds a carried
  // paragraph-relative drawing (WORD_LATER_ANCHOR_EARLIER_LINE_WRAP).
  // WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH carries likewise change their page.
  const carryBounded = (limit: (pass: BodyPaginationPassResult) => number) =>
    (pass: BodyPaginationPassResult) => Math.min(
      limit(pass),
      firstHostCarryPage(pass.layout),
      firstFrameCarryPage(pass),
    );
  // Every anchor pass may publish, but only the leading pages
  // `anchorStablePageLimit` proves every later pass of this run reproduces. A
  // seeded run never publishes: it can be abandoned for an unseeded retry,
  // whose passes form a different chain.
  const anchorPassObserver = (
    appliedPlan: ReadonlyMap<string, PageWrapDestination> | null,
    appliedPlanComplete: () => boolean = () => true,
  ): BodyPaginationPassObserver | undefined => {
    if (!publisher || seedPlan) return undefined;
    const events: PageAnchorInputEvent[] = [];
    // A prescan disagreement on a closed page is permanent for this pass.
    let ceiling = Number.POSITIVE_INFINITY;
    return passPublicationObserver(publisher, carryBounded((pass) => {
      const stable = anchorStablePageLimit(events, appliedPlan, pass, appliedPlanComplete());
      ceiling = Math.min(ceiling, stable.prescanDisagreementPage);
      return stable.limit;
    }), {
      onPageAnchorInput: (event) => { events.push(event); },
      canExtend: () => ceiling > publisher.publishedPages,
    });
  };
  /**
   * The unseeded run's first pass, with its second pass advanced behind it.
   *
   * The second pass applies the plan the first one observes. It reads that
   * plan only where it opens a page or flow domain (`pageStartAnchors`), and
   * only that page's destinations. A page the first pass has closed keeps its
   * destinations: tables are placed at or after the page that reaches them and
   * drawings on their anchor paragraph's page, so neither lands on a closed
   * page. The second pass may therefore lay out every page whose opening reads
   * fall on pages the first pass has closed, with the result identical to the
   * same pass started after the first one finishes. Its proven-stable prefix
   * (`anchorStablePageLimit`) can then be published while the first pass is
   * still discovering destinations, instead of only after it, which for a
   * document that opens with a page-positioned table is the whole first pass.
   *
   * The plan view grows in page order, as `pageAnchorDestinationPlan` builds
   * the full plan, from each page draft the first pass closes (no snapshot of
   * the pass is composed, so discovery stays linear in the page count). A
   * second-pass read of a page not yet closed suspends that pass; the driver
   * then steps the first pass, every step a scheduler suspension point, until
   * the page closes. When the first pass finishes, the next solver step must
   * apply exactly the view's plan, entry for entry; otherwise the overlapped
   * pass is discarded and the ordinary second pass runs instead. Final
   * output never depends on whether the overlap was used, only publication
   * timing does.
   */
  function* paginateOpeningPassWithLeadingSecond(): Generator<number, Readonly<{
    first: BodyPaginationPassResult;
    second: LeadingSecondPass | null;
  }>, void> {
    const view = new Map<string, PageWrapDestination>();
    // Pages below this index are closed in the first pass.
    let closedBefore = 0;
    let abandoned = false;
    const addDestinations = (pages: readonly LayoutPage[]) => {
      for (const [key, destination] of pageAnchorDestinationPlan({ pages, diagnostics: [] })) {
        // A destination on a closed page cannot reappear on a later one; if
        // it ever did, the view would not be the plan, so stop overlapping.
        if (view.has(key)) {
          abandoned = true;
          return;
        }
        view.set(key, destination);
      }
    };
    let firstResult: BodyPaginationPassResult | null = null;
    let secondResult: BodyPaginationPassResult | null = null;
    let secondLive = -1;
    // The page whose plan the suspended second pass is waiting for.
    let secondWaitsFor: number | null = null;
    const firstPublisher = anchorPassObserver(null)!;
    const firstObserver: BodyPaginationPassObserver = {
      shouldPublish: (committedPages) => firstPublisher.shouldPublish(committedPages),
      publish: (pass, processedEntries) => firstPublisher.publish(pass, processedEntries),
      onPageAnchorInput: (event) => firstPublisher.onPageAnchorInput?.(event),
      // Destinations come from each newly closed page alone: they depend
      // only on that page's body nodes and section regions, not on page
      // numbering or later pages, so no snapshot of the pass is composed.
      onPageDrafts: (pages) => {
        if (abandoned) return;
        const closed = pages.length - 1;
        if (closed <= closedBefore) return;
        const finalized: LayoutPage[] = [];
        for (let index = closedBefore; index < closed; index += 1) {
          const draft = pages[index]!;
          if (draft.kind === 'parity-blank') continue;
          // Section-owned identity is fixed once the previous page is (a
          // parity blank counts, §17.6.10); resolve it as `finalize` does so
          // a bordered page is built from the same facts.
          const firstSectionOwnedPage = index === 0
            || pages[index - 1]!.accumulator.sectionOccurrenceId !== draft.accumulator.sectionOccurrenceId;
          finalized.push(finalizeLayoutPage(draft.accumulator, CLOSED_PAGE_NUMBER, firstSectionOwnedPage));
        }
        addDestinations(finalized);
        closedBefore = closed;
      },
    };
    const first = paginateBodyPassSteps(
      input, services, options, reserves, null, null, balancePlan, firstObserver, null,
    );
    const secondPublisher = anchorPassObserver(view, () => firstResult !== null)!;
    const secondObserver: BodyPaginationPassObserver = {
      shouldPublish: (committedPages) => !abandoned && secondPublisher.shouldPublish(committedPages),
      publish: (pass, processedEntries) => secondPublisher.publish(pass, processedEntries),
      // One body entry can open several pages. The second pass suspends
      // until the first has closed the page this read needs.
      pageAnchorPlanPending: (pageIndex) => {
        secondWaitsFor = !abandoned && firstResult === null && closedBefore <= pageIndex
          ? pageIndex
          : null;
        return secondWaitsFor !== null;
      },
      onPageAnchorInput: (event) => secondPublisher.onPageAnchorInput?.(event),
    };
    const second = paginateBodyPassSteps(
      input, services, options, reserves, view, new Map(), balancePlan, secondObserver, null,
    );
    // Both passes advance one suspension at a time, each forwarded to the
    // scheduler, so host turns and cancellation are never starved.
    const secondMayStep = () => !abandoned && secondResult === null && (
      secondWaitsFor !== null
        ? closedBefore > secondWaitsFor
        // Keep the second pass at most one page behind the closed prefix.
        : closedBefore > secondLive + 1);
    while (firstResult === null) {
      if (secondMayStep()) {
        const next = second.next();
        if (next.done) secondResult = next.value;
        else {
          if (secondWaitsFor === null) secondLive = next.value - 1;
          yield next.value;
        }
        continue;
      }
      const step = first.next();
      if (step.done) {
        firstResult = step.value;
        if (!abandoned) {
          addDestinations(step.value.layout.pages.filter((page) => page.pageIndex >= closedBefore));
        }
        break;
      }
      yield step.value;
    }
    const firstPass = firstResult as BodyPaginationPassResult | null;
    if (firstPass === null) throw new Error('The opening anchor pass did not complete');
    if (abandoned) return Object.freeze({ first: firstPass, second: null });
    const appliedPlan = JSON.stringify([...view]);
    return Object.freeze({
      first: firstPass,
      second: Object.freeze({
        appliesCarry: (
          plan: ReadonlyMap<string, PageWrapDestination>,
          minimumTablePageBySource: ReadonlyMap<string, number>,
          deferrals: AnchorLineDeferrals,
        ) => (
          minimumTablePageBySource.size === 0 && deferrals.size === 0
          && JSON.stringify([...plan]) === appliedPlan
        ),
        finish: function* () {
          if (abandoned) throw new Error('An abandoned overlapped anchor pass cannot finish');
          return secondResult ?? (yield* second);
        },
      }),
    });
  }
  // The plan the most recent converge step observed; at a fixed point it is
  // the plan its returned pass applied.
  let convergedPlan: ReadonlyMap<string, PageWrapDestination> = new Map();
  const converge = function* (
    // `null` throughout, matching the solver's first-pass `previous`: an
    // unseeded run has no applied plan, which is not an omitted argument.
    initialPlan: ReadonlyMap<string, PageWrapDestination> | null,
    publish: boolean,
  ) {
    type AnchorPassCarry = Readonly<{
      plan: ReadonlyMap<string, PageWrapDestination>;
      minimumTablePageBySource: ReadonlyMap<string, number>;
      deferrals: AnchorLineDeferrals;
    }>;
    const noDeferrals: AnchorLineDeferrals = new Map();
    const settledFloors = new Map<string, number>();
    let restart: AnchorPassCarry | null = null;
    const latest: { carry: AnchorPassCarry | null } = { carry: null };
    let remainingPasses = anchorPassLimit;
    // Only the unseeded publishing run publishes, so only it overlaps its
    // second pass with its first (see `LeadingSecondPass`).
    let leadingSecond: LeadingSecondPass | null = null;
    const identity = (value: AnchorPassCarry) => anchorCarryIdentity(value.plan, value.deferrals)
      // Preserve main's detector until recovery. Once floors are pinned, they
      // are persistent solver input and must participate in exact equality.
      + (settledFloors.size > 0 ? `\u0005${anchorPlanIdentity(value.minimumTablePageBySource)}` : '');
    for (;;) {
      try {
        return (yield* convergeExactStateSteps<AnchorPassCarry & Readonly<{
          pass: BodyPaginationPassResult;
        }>, number, AnchorPassCarry>({
          ...(restart
            ? { seedState: identity(restart) }
            : initialPlan !== null ? { seedState: anchorCarryIdentity(initialPlan, noDeferrals) } : {}),
          step: function* anchorPass(carried) {
            const previous = carried ?? restart;
            const appliedPlan = previous !== null ? previous.plan : initialPlan;
            const appliedDeferrals = previous?.deferrals ?? noDeferrals;
            const overlapsSecond = publish && publisher !== undefined && seedPlan === undefined
              && previous === null && appliedPlan === null;
            const leading = leadingSecond;
            leadingSecond = null;
            const pass = overlapsSecond
              ? yield* (function* () {
                const opening = yield* paginateOpeningPassWithLeadingSecond();
                leadingSecond = opening.second;
                return opening.first;
              })()
              : leading !== null && previous !== null
                && leading.appliesCarry(previous.plan, previous.minimumTablePageBySource, appliedDeferrals)
                ? yield* leading.finish()
                : yield* paginateBodyPassSteps(
                  input,
                  services,
                  options,
                  reserves,
                  appliedPlan,
                  previous?.minimumTablePageBySource ?? null,
                  balancePlan,
                  publish ? anchorPassObserver(appliedPlan) : undefined,
                  appliedDeferrals.size > 0 ? appliedDeferrals : null,
                );
            const observed = pageAnchorDestinationPlan(pass.layout);
            for (const [key, destination] of keptHostAnchorCarries(appliedPlan, pass.layout)) {
              observed.set(key, destination);
            }
            for (const [key, destination] of keptFrameCarries(appliedPlan, pass)) {
              observed.set(key, destination);
            }
            // The unseeded pass registers source-order estimates, not anchor
            // lines that reached their page, so it proves nothing about them.
            const lineTests = appliedPlan === null
              ? Object.freeze({ plan: observed, deferrals: appliedDeferrals })
              : resolveAnchorLineTests(pass, observed, appliedDeferrals);
            const minimumTablePageBySource = new Map<string, number>();
            // Compare all destinations once. Testing every table against a fresh
            // copy of the full plan would make table-heavy documents quadratic.
            // A table floor needs the next pass to differ from this one only in
            // that table's registration, so no drawing retest or new deferral.
            const changedKeys = previous && appliedPlan
              && lineTests.plan === observed && lineTests.deferrals === appliedDeferrals
              ? changedAnchorKeys(appliedPlan, observed)
              : null;
            for (const [key, destination] of observed) {
              if (destination.kind !== 'floating-table') continue;
              const prior = appliedPlan?.get(key);
              if (prior?.kind !== 'floating-table' || !changedKeys
                || (changedKeys.size > 1 || (changedKeys.size === 1 && !changedKeys.has(key)))) {
                continue;
              }
              // Only this table's changed exclusion can have moved its source in
              // this run: input, reserves, and every other page anchor are fixed.
              // Recheck the next candidate page, not the observed source page.
              const provenPage = destination.pageIndex > prior.pageIndex
                ? prior.pageIndex + 1
                : previous?.minimumTablePageBySource.get(key);
              if (provenPage !== undefined) minimumTablePageBySource.set(key, provenPage);
            }
            // Only cycle recovery retains floors across passes. Ordinary runs
            // keep the original single-table proof and its exact output unchanged.
            for (const [key, pageIndex] of settledFloors) {
              const floor = Math.max(pageIndex, minimumTablePageBySource.get(key) ?? pageIndex);
              settledFloors.set(key, floor);
              minimumTablePageBySource.set(key, floor);
            }
            convergedPlan = lineTests.plan;
            latest.carry = Object.freeze({
              plan: lineTests.plan,
              minimumTablePageBySource,
              deferrals: lineTests.deferrals,
            });
            return Object.freeze({
              pass,
              plan: lineTests.plan,
              minimumTablePageBySource,
              deferrals: lineTests.deferrals,
            });
          },
          stateOf: identity,
          // The next anchor pass reads only the plan, the proven table pages and
          // the proven anchor-line deferrals; the superseded pass (its whole
          // layout) is not carried into it.
          carry: (value) => Object.freeze({
            plan: value.plan,
            minimumTablePageBySource: value.minimumTablePageBySource,
            deferrals: value.deferrals,
          }),
          limit: remainingPasses,
        })).value.pass;
      } catch (error) {
        // Keep the existing unseeded retry for a carried plan and the existing
        // host-drawing fallback. Only an otherwise-fatal source-order cycle
        // enters recovery, so successful retry paths retain their output.
        leadingSecond = null;
        if (initialPlan !== null || !publish || !(error instanceof ExactConvergenceError) || error.reason !== 'cycle'
          || latest.carry === null) throw error;
        // §17.4.57 defines page-relative geometry but not this solver's cycle
        // recovery. `word-page-table-later-yields`: pin only the last
        // oscillating table (in source order) to the latest page it reached
        // in the repeated cycle, and let the earlier tables settle around it;
        // a persisting cycle pins the next one on the following recovery.
        // Never return a pass with stale exclusions: normal acquisition then
        // rebuilds text and table placement with these floors.
        const repeated = error.states.at(-1)!;
        const cycle = error.states.slice(error.states.indexOf(repeated), -1)
          .map((state) => new Map(JSON.parse(state.split('\u0004')[0]!) as Array<[string, PageWrapDestination]>));
        const changed = new Set<string>();
        const later = new Map<string, number>();
        for (const plan of cycle) {
          for (const key of changedAnchorKeys(cycle[0]!, plan)) changed.add(key);
          for (const [key, destination] of plan) {
            if (destination.kind !== 'floating-table') continue;
            later.set(key, Math.max(later.get(key) ?? -1, destination.pageIndex));
          }
        }
        const pinned = [...later]
          .filter(([key, pageIndex]) => changed.has(key) && pageIndex > (settledFloors.get(key) ?? -1))
          .sort(([left], [right]) => compareTableSourceOrder(right, left))[0];
        if (!pinned) throw error;
        settledFloors.set(pinned[0], pinned[1]);
        // Every settlement strictly raises a table floor: at most one raise
        // per table per page. Restarts share the original operational pass
        // budget as well, so even a growing page count cannot run unbounded.
        remainingPasses -= error.passes;
        if (remainingPasses < 2) {
          throw new ExactConvergenceError('limit', error.states, anchorPassLimit);
        }
        const floors = new Map(latest.carry.minimumTablePageBySource);
        for (const [key, pageIndex] of settledFloors) {
          floors.set(key, Math.max(pageIndex, floors.get(key) ?? pageIndex));
        }
        restart = Object.freeze({ ...latest.carry, minimumTablePageBySource: floors });
      }
    }
  };
  const carryFree = function* () {
    if (pageOwnedAnchorCount === 0) {
      return yield* paginateBodyPassSteps(
        input, services, options, reserves, null, null, balancePlan,
        publisher ? passPublicationObserver(publisher, carryBounded(livePageIndex)) : undefined,
      );
    }
    try {
      try {
        return yield* converge(seedPlan ?? null, true);
      } catch (error) {
        if (!seedPlan || !(error instanceof ExactConvergenceError)) throw error;
        // A plan carried from another reserve/balance run is only a starting
        // estimate. Its former page ownership may be invalid in this run; retry
        // once from the unseeded source order before reporting non-convergence.
        return yield* converge(null, true);
      }
    } catch (error) {
      if (error instanceof ExactConvergenceError) {
        throw new LayoutInvariantError(
          'NON_CONVERGENCE',
          error.reason === 'cycle'
            ? 'Page-anchor destination acquisition repeated an exact-state cycle'
            : `Page-anchor destination acquisition reached the operational pass limit ${anchorPassLimit}`,
        );
      }
      throw error;
    }
  };
  const settled = yield* carryFree();
  if (settled.terminalDiagnostic !== null) return settled;
  const settledPlan = convergedPlan;
  // WORD_LATER_ANCHOR_EARLIER_LINE_WRAP: carry each qualifying
  // paragraph-relative drawing, with the geometry of its first placement, to
  // the start of its page so earlier lines wrap around it. A carried drawing
  // whose paragraph then lands elsewhere is dropped, so the carried set only
  // shrinks and the run terminates.
  const hostPlan = hostAnchorCarryPlan(settled.layout);
  let carried = settled;
  if (hostPlan.size > 0) {
    const carriedPlan = new Map<string, PageWrapDestination>(pageAnchorDestinationPlan(settled.layout));
    for (const [key, destination] of hostPlan) carriedPlan.set(key, destination);
    anchorPassLimit = ANCHOR_PASS_BASE_LIMIT + 4 * (pageOwnedAnchorCount + hostPlan.size);
    try {
      carried = yield* converge(carriedPlan, false);
    } catch (error) {
      if (!(error instanceof ExactConvergenceError)) throw error;
      convergedPlan = settledPlan;
      carried = settled;
    }
  }
  // WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: frames are carried one pusher per
  // frame run and pass, in source order, because each push moves the text the
  // later frames of that run are placed from. A frame that once lost its carry
  // is never carried again, so every round adds a frame group and the rounds
  // are bounded by the frame groups of the document.
  const everCarried = new Set<string>();
  let applied = convergedPlan;
  // Every round adds a plan entry (a carry or a frame-run deferral) that is
  // never re-added; the cap is an operational guard against an unbounded run
  // of deferrals, reported like the anchor solver's own pass limit.
  const roundLimit = ANCHOR_PASS_BASE_LIMIT + 4 * carried.frameObservations.length;
  for (let round = 0; ; round += 1) {
    const next = frameCarryPlan(carried, applied, everCarried);
    if (next === null) return carried;
    if (round >= roundLimit) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `Text-frame carry reached the operational round limit ${roundLimit}`,
      );
    }
    const frameEntries = [...next.values()]
      .filter((destination) => destination.kind === 'frame' || destination.kind === 'frame-deferral').length;
    anchorPassLimit = ANCHOR_PASS_BASE_LIMIT
      + 4 * (pageOwnedAnchorCount + hostPlan.size + frameEntries);
    try {
      carried = yield* converge(next, false);
      applied = convergedPlan;
    } catch (error) {
      if (!(error instanceof ExactConvergenceError)) throw error;
      // Never return a pass without a push it requires.
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        error.reason === 'cycle'
          ? 'Text-frame carry repeated an exact-state cycle'
          : `Text-frame carry reached the operational pass limit ${anchorPassLimit}`,
      );
    }
  }
}

/** The carried frames of `applied` that were placed at their carried
 * position on their carried page and flow domain in `pass`. */
function keptFrameCarries(
  applied: ReadonlyMap<string, PageWrapDestination> | null,
  pass: BodyPaginationPassResult,
): Map<string, PageWrapDestination> {
  const kept = new Map<string, PageWrapDestination>();
  if (!applied) return kept;
  for (const [key, destination] of applied) {
    if (destination.kind === 'frame-deferral') {
      kept.set(key, destination);
      continue;
    }
    if (destination.kind !== 'frame') continue;
    const observed = pass.frameObservations.find((observation) => observation.carry.key === key);
    if (observed?.carry.frozen === true
      && observed.pageIndex === destination.pageIndex
      && observed.flowDomainId === destination.flowDomainId) {
      kept.set(key, destination);
    }
  }
  return kept;
}

/** `applied` plus the first uncarried frame of each frame run whose exclusion
 * meets earlier content of its page in `pass`, or null when there is none. */
function frameCarryPlan(
  pass: BodyPaginationPassResult,
  applied: ReadonlyMap<string, PageWrapDestination>,
  everCarried: Set<string>,
): Map<string, PageWrapDestination> | null {
  const plan = new Map(applied);
  const kept = keptFrameCarries(applied, pass);
  for (const [key, destination] of applied) {
    if (destination.kind === 'frame' && !kept.has(key)) plan.delete(key);
  }
  const runs = new Set<string>();
  let added = false;
  for (const observation of pass.frameObservations) {
    const { carry } = observation;
    if (carry.frozen || !observation.meetsEarlierContent || runs.has(carry.runKey)) continue;
    runs.add(carry.runKey);
    if (everCarried.has(carry.key)) continue;
    everCarried.add(carry.key);
    plan.set(carry.key, Object.freeze({
      kind: 'frame',
      occurrenceId: carry.key,
      paragraphSource: carry.paragraphSource,
      pageIndex: observation.pageIndex,
      flowDomainId: observation.flowDomainId,
      carry: Object.freeze({ ...carry, frozen: false }),
    }));
    added = true;
  }
  // A frame run lands in the flow domain of its anchor paragraph: when the
  // anchor started in a later one, the frames move on from where they were.
  for (const observation of pass.frameAnchorObservations) {
    if (!observation.framesAfterContent
      || (observation.anchorPageIndex === observation.framesPageIndex
        && observation.anchorFlowDomainId === observation.framesFlowDomainId)) continue;
    const key = frameDeferralKey(
      observation.runKey, observation.framesPageIndex, observation.framesFlowDomainId,
    );
    if (plan.has(key)) continue;
    plan.set(key, Object.freeze({
      kind: 'frame-deferral',
      occurrenceId: key,
      runKey: observation.runKey,
      pageIndex: observation.framesPageIndex,
      flowDomainId: observation.framesFlowDomainId,
    }));
    added = true;
  }
  return added ? plan : null;
}

function frameDeferralKey(runKey: string, pageIndex: number, flowDomainId: string): string {
  return `frame-deferral|${runKey}|${pageIndex}|${flowDomainId}`;
}

/** Lowest page index holding a frame a carry would move text on. */
function firstFrameCarryPage(pass: BodyPaginationPassResult): number {
  let first = Number.POSITIVE_INFINITY;
  for (const observation of pass.frameObservations) {
    if (!observation.carry.frozen && observation.meetsEarlierContent) {
      first = Math.min(first, observation.pageIndex);
    }
  }
  return first;
}

function continuousBalanceBoundaries(input: BodyLayoutInput): readonly Readonly<{
  outgoingSectionOccurrenceId: string;
  incomingSectionOccurrenceId: string;
}>[] {
  const boundaries: Array<Readonly<{
    outgoingSectionOccurrenceId: string;
    incomingSectionOccurrenceId: string;
  }>> = [];
  let outgoing = input.initialSection;
  for (const entry of input.sequence) {
    if (entry.kind !== 'begin-section') continue;
    if (entry.section.startType === 'continuous') {
      boundaries.push(Object.freeze({
        outgoingSectionOccurrenceId: outgoing.sectionOccurrenceId,
        incomingSectionOccurrenceId: entry.section.sectionOccurrenceId,
      }));
    }
    outgoing = entry.section;
  }
  return Object.freeze(boundaries);
}

function sharedContinuousBoundaryPage(
  layout: DocumentLayout,
  outgoingSectionOccurrenceId: string,
  incomingSectionOccurrenceId: string,
) {
  for (const page of layout.pages) {
    for (let index = 0; index + 1 < page.sectionRegions.length; index += 1) {
      const outgoing = page.sectionRegions[index]!;
      const incoming = page.sectionRegions[index + 1]!;
      if (outgoing.sectionOccurrenceId === outgoingSectionOccurrenceId
        && incoming.sectionOccurrenceId === incomingSectionOccurrenceId) {
        return Object.freeze({ page, outgoing });
      }
    }
  }
  return null;
}

function* paginateBodyWithColumnBalancingSteps(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
  publisher?: BodyPagePublisher,
  seedPlan?: ReturnType<typeof pageAnchorDestinationPlan>,
): PaginationSteps<BodyPaginationPassResult> {
  let plan: BodyBalancePlan = new Map();
  let lastBodyEntry = input.sequence.length - 1;
  while (lastBodyEntry >= 0 && input.sequence[lastBodyEntry]!.kind !== 'body-block'
    && input.sequence[lastBodyEntry]!.kind !== 'adjacent-table-group') lastBodyEntry--;
  const terminalBodyPage = (layout: DocumentLayout) => {
    for (let index = layout.pages.length - 1; index >= 0; index--) {
      if (layout.pages[index]!.layers.body.length > 0) return layout.pages[index];
    }
    return undefined;
  };
  const seedPublisher = publisher && input.compatibilityMode === 15 && (input.endnoteIds?.length ?? 0) > 0
    ? {
        get publishedPages() { return publisher.publishedPages; },
        get failed() { return publisher.failed; },
        publish(pass: BodyPaginationPassResult, processedEntries: number, limit: number) {
          const terminal = processedEntries > lastBodyEntry
            ? terminalBodyPage(pass.layout)?.pageIndex : undefined;
          publisher.publish(pass, processedEntries, terminal === undefined ? limit : Math.min(limit, terminal));
        },
      } : publisher;
  let pass: BodyPaginationPassResult | null = yield* paginateBodyWithAnchorConvergenceSteps(
    input,
    services,
    options,
    reserves,
    plan,
    seedPublisher,
    seedPlan,
  );
  if (pass.terminalDiagnostic !== null) return pass;
  for (const boundary of continuousBalanceBoundaries(input)) {
    const target = continuousBalanceTarget(input, pass, boundary);
    if (target === null) continue;
    const nextPlan = new Map(plan);
    nextPlan.set(boundary.outgoingSectionOccurrenceId, target);
    plan = nextPlan;
    // The rebalanced pass reads only the balance plan and the accepted anchor
    // plan of this one. Release this pass before the next one builds so two
    // whole layouts are never live at once.
    const anchorPlan = pageAnchorDestinationPlan(pass.layout);
    pass = null;
    pass = yield* paginateBodyWithAnchorConvergenceSteps(
      input,
      services,
      options,
      reserves,
      plan,
      undefined,
      anchorPlan,
    );
    if (pass.terminalDiagnostic !== null) return pass;
  }
  if (input.compatibilityMode === 15 && pass.endnoteLayoutsByPage.size > 0) {
    const terminal = terminalBodyPage(pass.layout);
    const region = terminal?.sectionRegions.at(-1);
    if (terminal && region && region.flowDomainIds.length > 1) {
      const nextPlan = new Map(plan);
      nextPlan.set(region.sectionOccurrenceId, Object.freeze({ pageIndex: terminal.pageIndex,
        targetPt: exactRetainedColumnBalanceTarget(input, pass.allocations,
          pass.footnoteReserveByPage, terminal, region),
      }));
      const anchorPlan = pageAnchorDestinationPlan(pass.layout);
      pass = null;
      pass = yield* paginateBodyWithAnchorConvergenceSteps(input, services, options, reserves,
        nextPlan, undefined, anchorPlan);
    }
  }
  return pass;
}

/** The exact column-balance target one continuous boundary adds to the plan,
 * or null when the pass leaves nothing to balance there. */
function continuousBalanceTarget(
  input: BodyLayoutInput,
  pass: BodyPaginationPassResult,
  boundary: ReturnType<typeof continuousBalanceBoundaries>[number],
): Readonly<{ pageIndex: number; targetPt: number }> | null {
  const baseline = sharedContinuousBoundaryPage(
    pass.layout,
    boundary.outgoingSectionOccurrenceId,
    boundary.incomingSectionOccurrenceId,
  );
  if (baseline === null || baseline.outgoing.flowDomainIds.length < 2) return null;
  const pageIndex = baseline.page.pageIndex;
  const targetPt = exactRetainedColumnBalanceTarget(
    input,
    pass.allocations,
    pass.footnoteReserveByPage,
    baseline.page,
    baseline.outgoing,
  );
  return Object.freeze({ pageIndex, targetPt });
}

/** Compose one retained pass through the authoritative layout-to-paint
 * boundary. Progressive snapshots carry only endnote occurrences whose
 * physical owner has already been fixed by the pagination pass. */
function composeBodyPaginationResult(
  pass: BodyPaginationPassResult,
  input: BodyLayoutInput,
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
  options: LayoutOptions,
  includeDocumentEndnotes: boolean,
): DocumentLayout {
  const bodyComposed = composeCanonicalSectionFlow(
    pass.layout,
    pass.session,
    pass.allocations,
  );
  const noteLayoutSettings = input.noteLayoutSettings ?? Object.freeze({
    footnotePosition: 'pageBottom',
    endnotePosition: 'docEnd',
  });
  const pageStories = composePageStories(
    bodyComposed,
    pass.session,
    owners,
    pass.footnoteLayoutsByPage,
    pass.endnoteLayoutsByPage,
  );
  const hasRetainedFootnotes = pageStories.pages.some((page) =>
    page.layers.notes.some((note) => note.source.story === 'footnote'));
  const composed = hasRetainedFootnotes && noteLayoutSettings.footnotePosition !== 'pageBottom'
    ? appendUnsupportedNotePositionDiagnostic(
        pageStories,
        'footnote',
        noteLayoutSettings.footnotePosition,
        'pageBottom',
      )
    : pageStories;
  const authoredEndnoteIds = includeDocumentEndnotes
    ? [...pass.endnoteLayoutsByPage.values()].flatMap(notes => notes.map(note => note.source.storyInstance)) : [];
  const withEndnotes = authoredEndnoteIds.length > 0
    && noteLayoutSettings.endnotePosition !== 'docEnd'
    ? appendUnsupportedNotePositionDiagnostic(
        composed,
        'endnote',
        noteLayoutSettings.endnotePosition,
        'docEnd',
      )
    : composed;
  const sourceDiagnostics = [
    ...(input.parserDiagnostics ?? []),
    ...(pass.terminalDiagnostic === null ? [] : [pass.terminalDiagnostic]),
  ];
  const withParserDiagnostics = sourceDiagnostics.length === 0
    ? withEndnotes
    : Object.freeze({
        ...withEndnotes,
        diagnostics: Object.freeze([
          ...sourceDiagnostics,
          ...withEndnotes.diagnostics,
        ]),
      });
  const finalized = Object.freeze({
    ...withParserDiagnostics,
    pages: Object.freeze(withParserDiagnostics.pages.map(finalizePageBookmarkStarts)),
  });
  const withChangeBars = options.showTrackedChanges === true
    ? attachTrackChangeBars(finalized)
    : finalized;
  return withChangeBars as DocumentLayout;
}

/**
 * Lay out the whole body, suspendable between body entries.
 *
 * `paginateBody` drives this to completion synchronously; the async driver in
 * `body-paginator-async.ts` spreads it across event-loop turns. Both run this
 * one implementation, so a progressive layout cannot diverge from a blocking
 * one.
 */
export function* paginateBodySteps(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  observer?: BodyPaginationObserver,
): PaginationSteps<DocumentLayout> {
  // Every convergence pass and its field-acquisition service views share this
  // private memo, while a later variant/document pagination starts fresh.
  services = createParagraphAcquisitionCacheServicesView(services);
  const owners = ownerMap(input);
  let publishedPages = 0;
  let publicationFailed = false;
  const publisher: BodyPagePublisher | undefined = observer
    ? {
        get publishedPages() { return publishedPages; },
        get failed() { return publicationFailed; },
        publish: (pass, processedEntries, pageIndexLimit) => {
          // The live page still owns the transition edge: section-region and
          // page-final composition can change when the following page opens.
          // It is never published, and the pass-level rule may bound the
          // prefix further (see `anchorStablePageLimit`).
          const limit = Math.min(pageIndexLimit, livePageIndex(pass));
          const count = pass.layout.pages.findIndex((page) => page.pageIndex >= limit);
          if (count <= publishedPages) return;
          try {
            const composed = composeBodyPaginationResult(pass, input, owners, options, false);
            // This publication is deliberately provisional: header/footer and
            // pagination-field convergence may replace it, and consumers
            // receive `exact:false`.
            const publishable = Object.freeze({
              ...composed,
              pages: Object.freeze(composed.pages.slice(0, count)),
            }) as DocumentLayout;
            observer.onPages(publishable, processedEntries);
            publishedPages = count;
          } catch {
            // A provisional snapshot is best-effort. The same live pagination
            // session remains authoritative and must be allowed to finish.
            publicationFailed = true;
          }
        },
      }
    : undefined;
  let seed: BodyPaginationPassResult | null = yield* paginateBodyWithColumnBalancingSteps(
    input, services, options, [], publisher,
  );
  const convergence = convergeHeaderFooterReserveSteps<
    BodyPaginationPassResult,
    number,
    ReserveRepaginationCarry
  >(seed, {
    measure: (pass) => headerFooterReserves(pass),
    // A repagination reads only these facts of the pass before it (besides
    // the measured reserves). Carrying them, rather than the pass, lets that
    // pass's whole layout and pagination session be collected while the next
    // one is built.
    carry: (pass) => Object.freeze({
      fieldContexts: paginationFieldPageContexts(pass.layout),
      // Page-owned tables need one geometry-discovery pass and one exclusion
      // pass. Reuse the accepted geometry for this later reserve iteration
      // when stable; changed ownership still runs exact convergence anew.
      anchorPlan: pageAnchorDestinationPlan(pass.layout),
    }),
    repaginate: function* reserveRepagination(reserves, carried) {
      const contexts = carried.fieldContexts;
      const iterationServices = createFieldAcquisitionServicesView(services, {
        totalPages: contexts.length,
        resolveDestinationPage: (pageIndex) => contexts[pageIndex],
      });
      let observed = reserves;
      const openingObservations = new Set<string>();
      // Retry the same algorithm with actual opening observations. Bound retries
      // by the page ceiling, retain only recent exact identities, and share the
      // original hard miss ledger. Completed-pass exact reserve, field and anchor
      // convergence remains unchanged. A fixed small retry cap would reject
      // documents with many independent first/default section boundaries.
      for (let refinement = 0; refinement < MAX_BODY_LAYOUT_PAGES; refinement++) {
        const fingerprint = JSON.stringify(observed);
        if (openingObservations.has(fingerprint)) throw new LayoutInvariantError('NON_CONVERGENCE', 'Opened page reserve observations did not converge');
        openingObservations.add(fingerprint);
        if (openingObservations.size > MAX_OPENED_RESERVE_IDENTITIES) {
          openingObservations.delete(openingObservations.values().next().value!);
        }
        try {
          return yield* paginateBodyWithColumnBalancingSteps(
            input, iterationServices, options, observed, undefined, carried.anchorPlan,
          );
        } catch (error) {
          if (!(error instanceof OpenedPageReserveObservation)) throw error;
          observed = error.reserves;
        }
      }
      throw new LayoutInvariantError('NON_CONVERGENCE', 'Opened page reserve observation refinement limit exceeded');
    },
    identity: (pass) => paginationFieldPageContexts(pass.layout),
    requiresConvergence: seed.session.hasPaginationFields,
  });
  seed = null;
  const converged = (yield* convergence).result;
  return (yield* assertAndDeepFreezeDocumentLayoutSteps(
    composeBodyPaginationResult(converged, input, owners, options, true),
  )) as DocumentLayout;
}

export function paginateBody(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
): DocumentLayout {
  return drainPagination(paginateBodySteps(input, services, options));
}
