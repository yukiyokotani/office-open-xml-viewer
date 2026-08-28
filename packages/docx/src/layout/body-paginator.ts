import type { SectionLayoutContext } from '../layout-context.js';
import type {
  BodyLayoutInput,
  BodyParagraphSourceInput,
  BodySectionLayoutInput,
  BodyTableSourceInput,
} from './body-layout-input.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutSession,
  BodyTableContinuationCursor,
} from './body-layout-kernel.js';
import { NoteCapacityExceededError } from './body-layout-kernel.js';
import {
  commitPageFlowTransition,
  createBodyPaginationState,
  createCanonicalPageDraft,
  addPageFootnoteReserve,
  setBodyBalanceTarget,
  type BodyPageTransitionFactory,
  type BodyPaginationState,
  type CanonicalPageDraft,
} from './body-pagination.js';
import { assertAndDeepFreezeDocumentLayout } from './invariants.js';
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
  type PageFlowState,
} from './paginator.js';
import {
  createPageFlowSectionContext,
  physicalSectionGeometry,
  resolveSectionContextForPage,
} from './context.js';
import {
  transformRect,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
} from './coordinate-space.js';
import { selectParagraphFragment, type ParagraphFragmentCursor } from './paragraph-pagination.js';
import { paragraphGapAdjustment } from './paragraph-spacing.js';
import {
  endnoteIdsInRetainedSlice,
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
  wordActiveColumnBreakIndexes,
  wordEmptyKeepNextBridgesSuccessor,
  wordFlowNeutralPreBreakAnchorParagraph,
  wordPreBreakHostAnchorExtentPt,
  wordPreBreakInlineDrawingResource,
} from './body-pagination-compatibility.js';
import {
  wordContinuousSectionRestartDisplayNumber,
  wordTrailingEmptyMarkAdmissionAllowancePt,
} from './section-compatibility.js';
import { bodyOccurrenceKey } from './source-key.js';
import {
  convergeHeaderFooterReserveSteps,
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
  LayoutServices,
  NoteLayout,
  PaintNode,
  ParagraphLayout,
  SourceRef,
  TableLayout,
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
    writingMode: writingModeFromTextDirection(context.textDirection),
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
  const writingMode = writingModeFromTextDirection(section.textDirection);
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
  reserves: readonly HeaderFooterReserve[],
): BodyPageTransitionFactory {
  const owner = (id: string) => {
    const result = owners.get(id);
    if (!result) throw new Error(`Unknown body section ${id}`);
    return result;
  };
  return {
    openContentPage(event) {
      const nextOwner = owner(event.sectionOccurrenceId);
      const reserve = reserves[event.pageIndex] ?? { top: 0, bottom: 0 };
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
        reserves[event.pageIndex] ?? { top: 0, bottom: 0 },
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
        // physical jc alignment followed by signed tblInd translation. Move
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

function tableCursorKey(cursor: import('./table-pagination.js').TableFragmentCursor): readonly unknown[] {
  return [
    cursor.rowIndex,
    cursor.rowFragmentIndex,
    cursor.cells.map((cell) => [
      cell.blockIndex,
      cell.paragraphLineStart,
      cell.nestedFragmentIndex,
      cell.nestedCursor === null ? null : tableCursorKey(cell.nestedCursor),
    ]),
  ];
}

function tableFragmentStartKey(cursor: BodyTableContinuationCursor | undefined): string {
  if (cursor === undefined) return 'root';
  if (cursor.kind === 'table') return `table:${JSON.stringify(tableCursorKey(cursor.cursor))}`;
  const tableCursor = cursor.cursor.tableCursor;
  return `adjacent-table:${cursor.cursor.tableIndex}:${cursor.cursor.sourceRowIndex}:${JSON.stringify(
    tableCursor === undefined ? null : tableCursorKey(tableCursor),
  )}`;
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

function finalize(state: BodyPaginationState, owners: ReadonlyMap<string, BodySectionLayoutInput>): DocumentLayout {
  // Compatibility-owned continuous-section restart arithmetic anchors the
  // incoming owner to the first physical page that retains its body content.
  // A transition-only empty region is capacity, not a numbering appearance.
  // Issue #804 locks the retained-layout and painted-footer observations together
  // in the continuous-section cases in page-number-field-render.test.ts.
  const contentFirstAppearance = sectionContentFirstAppearancePageIndices(
    state.pages.map((draft) => ({
      pageIndex: draft.accumulator.pageIndex,
      sectionRegions: draft.accumulator.sectionRegions.map((region) => ({
        sectionOccurrenceId: region.sectionOccurrenceId,
        flowDomainIds: (region.columnIndexes
          ?? region.section.columns.map((_, index) => index))
          .map((columnIndex) => bodyFlowDomainId(
            draft.accumulator.pageIndex,
            region.id,
            columnIndex,
          )),
      })),
      contentFlowDomainIds: draft.accumulator.readingOrder.map((node) => node.flowDomainId),
    })),
  );
  let displayNumber = 0;
  let priorOwner: string | null = null;
  const pages = state.pages.map((draft) => {
    const owner = owners.get(draft.accumulator.sectionOccurrenceId)!;
    const firstSectionOwnedPage = owner.sectionOccurrenceId !== priorOwner;
    if (owner.sectionOccurrenceId !== priorOwner && owner.pageNumbering.start !== null) {
      displayNumber = wordContinuousSectionRestartDisplayNumber(
        owner.pageNumbering.start,
        draft.accumulator.pageIndex,
        contentFirstAppearance.get(owner.sectionOccurrenceId)
          ?? draft.accumulator.pageIndex,
      );
    } else displayNumber += 1;
    priorOwner = owner.sectionOccurrenceId;
    const pageNumber = {
      displayNumber,
      format: owner.pageNumbering.format ?? 'decimal',
      sectionOccurrenceId: owner.sectionOccurrenceId,
    };
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
  return layout;
}

/** One completed body-pagination pass. */
export type BodyPaginationPassResult = Readonly<{
  layout: DocumentLayout;
  session: BodyLayoutSession;
  allocations: readonly BodyFlowAllocation[];
  footnoteReserveByPage: ReadonlyMap<number, number>;
  footnoteLayoutsByPage: ReadonlyMap<number, readonly NoteLayout[]>;
  terminalDiagnostic: LayoutDiagnostic | null;
}>;

interface BodyPaginationPassObserver {
  shouldPublish(committedPages: number): boolean;
  publish(pass: BodyPaginationPassResult, processedEntries: number): void;
}

/** Optional observer for paintable snapshots produced by the first canonical
 * pagination pass. The generator itself continues from the same suspended
 * state after every publication; no source prefix is replayed. */
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
  terminalDiagnostic: LayoutDiagnostic | null,
): BodyPaginationPassResult {
  const layout = finalize(state, owners);
  const retainedPageIndexes = new Set(layout.pages.map((page) => page.pageIndex));
  const retainedNodeIds = new Set(layout.pages.flatMap((page) => (
    pageLayerNodes(page).map(({ node }) => node.id)
  )));
  return Object.freeze({
    layout,
    session,
    allocations: Object.freeze(allocations.filter((allocation) => (
      retainedNodeIds.has(allocation.nodeId)
    ))),
    footnoteReserveByPage: new Map([...footnoteReserveByPage]
      .filter(([pageIndex]) => retainedPageIndexes.has(pageIndex))),
    footnoteLayoutsByPage: new Map([...footnoteLayoutsByPage]
      .filter(([pageIndex]) => retainedPageIndexes.has(pageIndex))),
    terminalDiagnostic,
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

function* paginateBodyPassSteps(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
  anchorDestinations: ReadonlyMap<string, Readonly<{
    occurrenceId: string;
    paragraphSource: SourceRef;
    pageIndex: number;
    flowDomainId: string;
  }>> | null,
  balancePlan: BodyBalancePlan,
  observer?: BodyPaginationPassObserver,
): PaginationSteps<BodyPaginationPassResult> {
  const kernel = bodyLayoutKernelOf(services);
  if (!kernel) throw new Error('Body layout kernel is not attached to the supplied services');
  const owners = ownerMap(input);
  const acceptedOccurrenceIds = new Set<string>();
  const allocations: BodyFlowAllocation[] = [];
  const initialReserve = reserves[0] ?? { top: 0, bottom: 0 };
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
  const factory = transitionFactory(owners, reserves);
  // Source entry whose keep-with-next set was relocated to a new physical page
  // by automatic overflow. The compatibility projection suppresses that
  // leading paragraph's space-before only for this grouped relocation;
  // ordinary overflow and authored page/section breaks retain their own rules.
  let automaticPageStartEntryIndex: number | null = null;
  const session = kernel.openBodyLayoutSession({
    source: input.source,
    section: input.initialSection.context,
    initialLocation: acquisitionLocation(state),
  }, services, options);
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
  const pageStartAnchors = (target: BodyPaginationState, startIndex: number) => {
    if (anchorDestinations !== null) {
      const location = acquisitionLocation(target);
      return Object.freeze([...anchorDestinations.values()]
        .filter((destination) => (
          destination.pageIndex === location.pageIndex
          && destination.flowDomainId === location.flowDomainId
        ))
        .map(({ occurrenceId, paragraphSource }) => Object.freeze({
          occurrenceId,
          paragraphSource,
        })));
    }
    const anchors: Array<Readonly<{ occurrenceId: string; paragraphSource: SourceRef }>> = [];
    for (let index = startIndex; index < input.sequence.length; index += 1) {
      const entry = input.sequence[index]!;
      if (entry.kind === 'authored-break' && entry.break !== 'column') break;
      if (entry.kind === 'begin-section' && entry.section.startType !== 'continuous') break;
      if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') continue;
      if (index > startIndex && entry.block.pageBreakBefore) break;
      entry.block.pageOwnedAnchorOccurrenceIds?.forEach((occurrenceId) => anchors.push(
        Object.freeze({ occurrenceId, paragraphSource: entry.block.source }),
      ));
    }
    return Object.freeze(anchors);
  };
  const prescanPageAnchors = (target: BodyPaginationState, startIndex: number) => {
    const anchors = pageStartAnchors(target, startIndex);
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
  prescanPageAnchors(state, 0);
  const commitTransition = (
    transition: ReturnType<typeof applyAuthoredBreak>,
    nextEntryIndex: number,
    suppressFirstParagraphSpaceBefore = false,
  ) => {
    const previousPageIndex = state.flow.pageIndex;
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
    const nextLocation = acquisitionLocation(state);
    if (state.flow.pageIndex !== previousPageIndex) {
      automaticPageStartEntryIndex = opensAutomaticPage && suppressFirstParagraphSpaceBefore
        ? nextEntryIndex
        : null;
      session.resetPageAcquisition(nextLocation);
      prescanPageAnchors(state, nextEntryIndex);
    } else {
      session.moveAcquisitionCursor(nextLocation);
      // §17.18.77 keeps the physical page but opens a distinct flow domain.
      // The outgoing source scan intentionally stopped at the section mark, so
      // acquire incoming page-owned wrap authority before its first paragraph.
      if (opensSamePageColumnRegion) {
        prescanPageAnchors(state, nextEntryIndex);
      }
    }
  };
  const footnoteIdsByPage = new Map<number, Set<string>>();
  const footnoteReserveByPage = new Map<number, number>();
  const footnoteLayoutsByPage = new Map<number, NoteLayout[]>();
  // §17.18.77 observes committed references even when their measured reserve is zero;
  // footnoteReservePt remains only the page-local geometry charge.
  const hasFootnoteReferenceOnPage = (pageIndex: number): boolean => (
    (footnoteIdsByPage.get(pageIndex)?.size ?? 0) > 0
  );
  const footnoteAdmission = (
    candidate: ParagraphLayout | TableLayout,
    inlineExtentPt: number,
    retainedReferenceIds?: readonly string[],
  ): Readonly<{
    ids: readonly string[];
    layouts: readonly NoteLayout[];
    reservePt: number;
  }> => footnoteAdmissionForIds(
    retainedReferenceIds ?? footnoteIdsInRetainedSlice(candidate),
    inlineExtentPt,
  );
  const footnoteAdmissionForIds = (
    retainedReferenceIds: readonly string[],
    inlineExtentPt: number,
  ): Readonly<{
    ids: readonly string[];
    layouts: readonly NoteLayout[];
    reservePt: number;
  }> => {
    const retained = footnoteIdsByPage.get(state.flow.pageIndex) ?? new Set<string>();
    const ids = [...new Set(retainedReferenceIds)]
      .filter((id) => !retained.has(id));
    const location = acquisitionLocation(state);
    if (ids.length > 0 && !session.layoutNotes) {
      throw new Error('Footnote layout requires a note-capable layout session');
    }
    const layouts = ids.length === 0 ? Object.freeze([]) : session.layoutNotes!({
      kind: 'footnote',
      referenceIds: Object.freeze(ids),
      pageIndex: state.flow.pageIndex,
      section: location.section,
      container: {
        id: `notes:page:${state.flow.pageIndex}`,
        kind: 'footnote',
        bounds: {
          xPt: location.availableBounds.xPt,
          yPt: 0,
          widthPt: inlineExtentPt,
          heightPt: location.section.geometry.pageHeight,
        },
      },
      firstOnPage: retained.size === 0,
    });
    return Object.freeze({
      ids: Object.freeze(ids),
      layouts,
      reservePt: layouts.reduce((sum, note) => sum + note.advancePt, 0),
    });
  };
  const commitFootnotes = (
    ids: readonly string[],
    layouts: readonly NoteLayout[],
    reservePt: number,
  ) => {
    let retained = footnoteIdsByPage.get(state.flow.pageIndex);
    if (!retained) {
      retained = new Set<string>();
      footnoteIdsByPage.set(state.flow.pageIndex, retained);
    }
    ids.forEach((id) => retained!.add(id));
    const retainedLayouts = footnoteLayoutsByPage.get(state.flow.pageIndex) ?? [];
    retainedLayouts.push(...layouts);
    footnoteLayoutsByPage.set(state.flow.pageIndex, retainedLayouts);
    footnoteReserveByPage.set(
      state.flow.pageIndex,
      (footnoteReserveByPage.get(state.flow.pageIndex) ?? 0) + reservePt,
    );
    state = addPageFootnoteReserve(state, reservePt);
  };
  // §17.11.21 / §17.18.34 assign each note to the physical page that paints
  // its reference. Growing that page-wide band must not clip a deeper column
  // that the immutable paginator has already committed.
  const additionalFootnoteReserveCapacityPt = (): number => Math.max(
    0,
    activeBlockEndPt(state)
      - state.footnoteReservePt
      - state.flow.deepestColumnBlockPt,
  );
  const footnoteReserveInvadesCommittedPageContent = (reservePt: number): boolean => (
    reservePt > additionalFootnoteReserveCapacityPt()
  );
  let previousParagraph: BodyParagraphSourceInput | null = null;
  const activeColumnBreakIndexes = wordActiveColumnBreakIndexes(input.sequence);
  let terminalDiagnostic: LayoutDiagnostic | null = null;

  bodyEntries: for (let entryIndex = 0; entryIndex < input.sequence.length; entryIndex += 1) {
    // Suspension point. `state.pages.length` is the committed page count, which
    // is what a progressive driver reports; a synchronous driver discards it.
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
        terminalDiagnostic,
      ), entryIndex);
    }
    const entry = input.sequence[entryIndex]!;
    if (entry.kind === 'consume-source') {
      continue;
    }
    if (entry.kind === 'authored-break') {
      previousParagraph = null;
      if (entry.break === 'column' && !activeColumnBreakIndexes.has(entryIndex)) {
        continue;
      }
      commitTransition(
        applyAuthoredBreak(state.flow, entry.break, entry.parity),
        entryIndex + 1,
      );
      continue;
    }
    if (entry.kind === 'begin-section') {
      previousParagraph = null;
      const currentWritingMode = writingModeFromTextDirection(activeRegion(state).section.textDirection);
      const incomingWritingMode = writingModeFromTextDirection(
        sectionContextForPage(entry.section, state.flow.pageIndex).textDirection,
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
      try {
        commitTransition(
          beginSection(
            state.flow,
            flowSection(entry.section, state.flow.pageIndex),
            effectiveStartType,
            {
              hasFootnoteReferenceOnCurrentPage: hasFootnoteReferenceOnPage(state.flow.pageIndex),
              incomingPageContentStartBlockPt: incomingInterval.blockStartPt,
              incomingPageContentEndBlockPt: incomingInterval.blockEndPt,
            },
          ),
          entryIndex + 1,
        );
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
      if (block.pageBreakBefore) {
        commitTransition(
          applyAuthoredBreak(state.flow, 'pageBreakBefore'),
          entryIndex,
        );
      }
      const previousAfterPt = previousParagraph?.spaceAfterPt ?? 0;
      const spacing = paragraphGapAdjustment(
        previousParagraph,
        block,
        previousAfterPt,
        block.continuousSectionRole === 'suppress-before' ? 0 : block.spaceBeforePt,
      );
      const spacingOverlap = block.continuousSectionRole === 'drop-previous-after'
        ? previousAfterPt
        : spacing.overlap;
      if (spacingOverlap > 0) {
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
      let cursor: ParagraphFragmentCursor | null = Object.freeze({ boundary: null });
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
              && !state.flow.pageHasContent
              && automaticPageStartEntryIndex === entryIndex
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
            && state.flow.pageHasContent
          ) {
            commitTransition(
              reserveInvadesCommittedPageContent
                ? advanceToPage(state.flow, state.flow.section, 'overflow')
                : advanceColumnOrPage(state.flow, 'overflow'),
              entryIndex,
            );
            continue;
          }
          // A placed frame still paints its retained references despite a zero flow
          // charge, so note ownership is committed with the accepted occurrence.
          state = acceptNode(
            state,
            acquired.layout,
            block.source,
            acquired.blockExtentPt,
            fragmentStartKey,
            acceptedOccurrenceIds,
            allocations,
            acquired.placement,
          );
          commitFootnotes(notes.ids, notes.layouts, notes.reservePt);
          if (acquired.flowRegistryDelta) {
            session.commitFlowRegistryDelta(acquired.flowRegistryDelta);
          }
          cursor = null;
          session.moveAcquisitionCursor(acquisitionLocation(state));
          continue;
        }
        if (cursor.boundary === null && block.keepNext && state.flow.pageHasContent) {
          let keepSetExtentPt = acquired.blockExtentPt;
          const keepSetReferenceIds = new Set(footnoteIdsInRetainedSlice(acquired.layout));
          let hasTerminalBlock = false;
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
            keepSetExtentPt += continues
              ? following.fullExtentPt
              : following.leadContentExtentPt;
            const referenceIds = continues
              ? following.fullFootnoteReferenceIds
              : following.leadFootnoteReferenceIds;
            referenceIds?.forEach((id) => keepSetReferenceIds.add(id));
            if (!continues) {
              hasTerminalBlock = true;
              break;
            }
          }
          const keepSetReservePt = footnoteAdmissionForIds(
            [...keepSetReferenceIds],
            location.availableBounds.widthPt,
          ).reservePt;
          const keepSetAdmissionPt = keepSetExtentPt + keepSetReservePt;
          if (
            hasTerminalBlock
            && keepSetAdmissionPt > location.availableBounds.heightPt
            && keepSetAdmissionPt <= freshPageExtent(state)
          ) {
            commitTransition(
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
              commitTransition(
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
          const neutral = wordFlowNeutralPreBreakAnchorParagraph(acquired.layout);
          if (neutral !== null) {
            state = acceptNode(
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
        const markReservePt = footnoteAdmission(
          acquired.layout,
          location.availableBounds.widthPt,
        ).reservePt;
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
          });
        const selected = selectParagraphFragment(
          acquired.layout,
          cursor,
          acquired.fragmentation,
          location.availableBounds.heightPt + trailingMarkAdmissionAllowancePt,
          freshPageExtent(state),
          state.flow.pageHasContent,
          {
            keepLines: block.keepLines,
            widowControl: block.widowControl,
            authoredSpaceAfterPt: block.spaceAfterPt,
            writingMode: activeRegion(state).writingMode,
          },
          (fragment) => footnoteAdmission(
            fragment,
            location.availableBounds.widthPt,
          ).reservePt,
          acquired.uniformRubyAdvancePt,
          (reservePt) => !footnoteReserveInvadesCommittedPageContent(reservePt),
        );
        if (selected.requiresFreshFlowRegion) {
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
          continue;
        }
        if (!selected.fragment) throw new Error('Paragraph acquisition made no progress');
        state = acceptNode(
          state,
          selected.fragment,
          block.source,
          Math.min(selected.admittedBlockExtentPt, location.availableBounds.heightPt),
          fragmentStartKey,
          acceptedOccurrenceIds,
          allocations,
          acquired.placement,
        );
        const notes = footnoteAdmission(
          selected.fragment,
          location.availableBounds.widthPt,
        );
        assertFreshPageFootnoteAdmission(
          notes.reservePt,
          selected.fragment.advancePt + notes.reservePt,
          freshPageExtent(state),
        );
        commitFootnotes(notes.ids, notes.layouts, notes.reservePt);
        if (acquired.flowRegistryDelta) {
          const acceptedDelta = paragraphFlowRegistryDeltaForAcceptedFragment(
            acquired.flowRegistryDelta,
            selected.fragment,
          );
          if (acceptedDelta) session.commitFlowRegistryDelta(acceptedDelta);
        }
        cursor = selected.nextCursor;
        if (cursor) {
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
        }
        location = acquisitionLocation(state);
        session.moveAcquisitionCursor(location);
      }
      previousParagraph = block;
    } else {
      previousParagraph = null;
      let cursor: import('./body-layout-kernel.js').BodyTableContinuationCursor | undefined;
      let complete = false;
      while (!complete) {
        const fragmentStartKey = tableFragmentStartKey(cursor);
        const location = acquisitionLocation(state);
        const requestAt = (availableBlockExtentPt: number) => session.measureTable({
            input: block,
            location,
            availableInlineExtentPt: location.availableBounds.widthPt,
            availableBlockExtentPt,
            freshPageBlockExtentPt: freshPageExtent(state),
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
        let lastFootnoteAdmission = Object.freeze({
          reservePt: notes.reservePt,
          chargePt: acquired.blockExtentPt + notes.reservePt,
        });
        const seenCandidates = new Set<string>();
        while (
          !acquired.requiresFreshFlowRegion
          && acquired.blockExtentPt + notes.reservePt > location.availableBounds.heightPt
        ) {
          const fingerprint = JSON.stringify({
            advancePt: acquired.blockExtentPt,
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
              chargePt: acquired.blockExtentPt + notes.reservePt,
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
            && acquired.nextCursor?.kind === 'table'
            && acquired.nextCursor.floatingContinuationFrame === 'fresh-text'
            && !(cursor?.kind === 'table' && cursor.floatingContinuationFrame !== undefined);
          if (acquired.nextCursor?.kind === 'table'
            && acquired.nextCursor.floatingContinuationFrame !== undefined) {
            cursor = acquired.nextCursor;
          }
          if (rebasesFloatingTableOnFreshFrame) continue;
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
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
          commitTransition(
            advanceToPage(state.flow, state.flow.section, 'overflow'),
            entryIndex,
          );
          continue;
        }
        state = acceptNode(
          state,
          acquired.layout,
          block.source,
          acquired.blockExtentPt,
          fragmentStartKey,
          acceptedOccurrenceIds,
          allocations,
          acquired.placement,
        );
        commitFootnotes(notes.ids, notes.layouts, notes.reservePt);
        if (acquired.flowRegistryDelta) {
          session.commitFlowRegistryDelta(bindTableFlowRegistryDeltaToAcceptedOccurrence(
            acquired.flowRegistryDelta,
            acquired.layout,
            bodyOccurrenceKey(block.source, location.flowDomainId, fragmentStartKey),
          ));
        }
        cursor = acquired.nextCursor ?? undefined;
        complete = cursor === undefined;
        if (cursor) {
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
        }
      }
    }
    session.moveAcquisitionCursor(acquisitionLocation(state));
  }
  const reservedPages = new Set([
    ...footnoteReserveByPage.keys(),
    ...footnoteLayoutsByPage.keys(),
  ]);
  for (const pageIndex of reservedPages) {
    const reservePt = footnoteReserveByPage.get(pageIndex) ?? 0;
    const retainedAdvancePt = (footnoteLayoutsByPage.get(pageIndex) ?? [])
      .reduce((sum, note) => sum + note.advancePt, 0);
    if (reservePt !== retainedAdvancePt) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `Page ${pageIndex} footnote reserve ${reservePt} does not equal retained advance ${retainedAdvancePt}`,
      );
    }
  }
  return paginationPassResult(
    state,
    owners,
    session,
    allocations,
    footnoteReserveByPage,
    footnoteLayoutsByPage,
    terminalDiagnostic,
  );
}

function headerFooterReserves(
  pass: Readonly<{ layout: DocumentLayout; session: BodyLayoutSession }>,
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
): readonly HeaderFooterReserve[] {
  return Object.freeze(pass.layout.pages.map((page, pageIndex) => {
    if (page.parityBlank) return Object.freeze({ top: 0, bottom: 0 });
    // Vertical header/footer stories paint in physical page space; charging their
    // measured overflow to the logical body interval would create a pagination-only reserve.
    if (writingModeFromTextDirection(page.section.textDirection) !== 'horizontal-tb') {
      return Object.freeze({ top: 0, bottom: 0 });
    }
    const owner = owners.get(page.sectionOccurrenceId);
    if (!owner) throw new Error(`Unknown body section ${page.sectionOccurrenceId}`);
    const inlineExtentPt = Math.max(
      0,
      page.section.geometry.pageWidth
        - Math.abs(page.section.geometry.marginLeft)
        - Math.abs(page.section.geometry.marginRight),
    );
    const measure = (kind: 'header' | 'footer') => {
      const source = selectedHeaderFooterStory(
        kind === 'header' ? owner.headers : owner.footers,
        {
          titlePage: owner.titlePage,
          firstPageOfSection: isFirstSectionOwnedPage(pass.layout.pages, pageIndex),
          evenAndOddHeaders: owner.evenAndOddHeaders,
          displayPageNumber: page.pageNumber.displayNumber,
        },
      );
      if (source === null) return 0;
      if (!pass.session.layoutStory) {
        throw new Error('Header/footer story layout requires a story-capable layout session');
      }
      return pass.session.layoutStory({
        source,
        pageIndex: page.pageIndex,
        section: page.section,
        container: {
          id: `story:${kind}:page:${page.pageIndex}`,
          kind,
          bounds: {
            xPt: Math.abs(page.section.geometry.marginLeft),
            yPt: 0,
            widthPt: inlineExtentPt,
            heightPt: page.section.geometry.pageHeight,
          },
        },
      }).advancePt;
    };
    return Object.freeze({
      top: headerFooterOverflowReservePt(
        measure('header'),
        page.section.geometry.marginTop,
        page.section.geometry.headerDistance,
      ),
      bottom: headerFooterOverflowReservePt(
        measure('footer'),
        page.section.geometry.marginBottom,
        page.section.geometry.footerDistance,
      ),
    });
  }));
}

function composePageStories(
  layout: DocumentLayout,
  session: BodyLayoutSession,
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
  footnotesByPage: ReadonlyMap<number, readonly NoteLayout[]>,
): DocumentLayout {
  const pages = layout.pages.map((page, pageIndex) => {
    if (page.parityBlank) return page;
    const owner = owners.get(page.sectionOccurrenceId);
    if (!owner) throw new Error(`Unknown body section ${page.sectionOccurrenceId}`);
    if (!session.layoutStory) {
      const hasPageStories = Object.values(owner.headers).some((source) => source !== null)
        || Object.values(owner.footers).some((source) => source !== null)
        || (footnotesByPage.get(page.pageIndex)?.length ?? 0) > 0;
      if (!hasPageStories) return page;
      throw new Error('Page-story composition requires a story-capable layout session');
    }
    const vertical = writingModeFromTextDirection(page.section.textDirection) !== 'horizontal-tb';
    const geometry = vertical
      ? physicalSectionGeometry(page.section.geometry)
      : page.section.geometry;
    const inlineStartPt = Math.abs(geometry.marginLeft);
    const inlineExtentPt = Math.max(
      0,
      geometry.pageWidth - Math.abs(geometry.marginLeft) - Math.abs(geometry.marginRight),
    );
    const coordinateSpace = vertical ? 'upright-physical' as const : 'section-logical' as const;
    const pageStorySection: DeepReadonly<SectionLayoutContext> = vertical
      ? Object.freeze({
          ...page.section,
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
      const story = session.layoutStory!({
        source,
        pageIndex: page.pageIndex,
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
      });
      const targetYPt = kind === 'header'
        ? geometry.headerDistance
        : geometry.pageHeight - geometry.footerDistance - story.advancePt;
      return translateStoryLayout(story, {
        xPt: 0,
        yPt: targetYPt - story.flowBounds.yPt,
      });
    };
    const header = acquire('header');
    const footer = acquire('footer');
    const retainedNotes = footnotesByPage.get(page.pageIndex) ?? [];
    const noteAdvancePt = retainedNotes.reduce((sum, note) => sum + note.advancePt, 0);
    const pageStoryRegion = page.sectionRegions[0];
    const noteBlockEndPt = pageStoryRegion?.blockEndPt
      ?? Math.max(
        0,
        page.section.geometry.pageHeight - Math.abs(page.section.geometry.marginBottom),
      );
    const noteTargetTopPt = noteBlockEndPt - noteAdvancePt;
    let noteCursorPt = noteTargetTopPt;
    const notes = retainedNotes.map((note) => {
      const translated = translateNoteLayout(note, {
        xPt: 0,
        yPt: noteCursorPt - note.flowBounds.yPt,
      });
      noteCursorPt += note.advancePt;
      return translated;
    });
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
    const notePhysicalBounds = pageStoryRegion
      ? Object.freeze(transformRect(
          pageStoryRegion.coordinateSpace.logicalToPhysical,
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
      ...(notes.length > 0 ? [Object.freeze({
        id: `notes:page:${page.pageIndex}`,
        kind: 'footnote' as const,
        ...(pageStoryRegion ? { sectionRegionId: pageStoryRegion.id } : {}),
        logicalBounds: noteLogicalBounds,
        physicalBounds: notePhysicalBounds,
      })] : []),
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
  return Object.freeze({ ...layout, pages: Object.freeze(pages) });
}

function composeDocumentEndnotes(
  layout: DocumentLayout,
  session: BodyLayoutSession,
  referenceIds: readonly string[],
): DocumentLayout {
  if (referenceIds.length === 0) return layout;
  let pageIndex = -1;
  for (let index = layout.pages.length - 1; index >= 0; index -= 1) {
    if (!layout.pages[index]!.parityBlank) {
      pageIndex = index;
      break;
    }
  }
  if (pageIndex < 0) return layout;
  const page = layout.pages[pageIndex]!;
  if (!session.layoutNotes) {
    return Object.freeze({
      ...layout,
      diagnostics: Object.freeze([...layout.diagnostics, Object.freeze({
        code: 'UNSUPPORTED_FEATURE' as const,
        severity: 'error' as const,
        source: Object.freeze({
          story: 'endnote' as const,
          storyInstance: referenceIds[0]!,
          path: Object.freeze([]),
        }),
        message: 'Document-end notes require a note-capable layout session',
      })]),
    });
  }
  const domains = new Map(page.flowDomains.map((domain) => [domain.id, domain]));
  const bodyNodes = page.layers.body.filter((node) => (
    node.ordinaryFlow && domains.get(node.flowDomainId)?.kind === 'body'
  ));
  const terminalBody = bodyNodes.reduce<PaintNode | null>((latest, node) => (
    latest === null
      || node.flowBounds.yPt + node.flowBounds.heightPt
        > latest.flowBounds.yPt + latest.flowBounds.heightPt
      ? node
      : latest
  ), null);
  const bodyDomain = terminalBody
    ? domains.get(terminalBody.flowDomainId)
    : [...page.flowDomains].reverse().find((domain) => domain.kind === 'body');
  if (!bodyDomain) {
    return Object.freeze({
      ...layout,
      diagnostics: Object.freeze([...layout.diagnostics, Object.freeze({
        code: 'UNSUPPORTED_FEATURE' as const,
        severity: 'error' as const,
        message: 'Document-end notes require a retained body flow domain',
      })]),
    });
  }
  const bodyRegion = page.sectionRegions.find((region) =>
    region.flowDomainIds.includes(bodyDomain.id));
  const endnoteRegion = bodyRegion ?? page.sectionRegions[0];
  const blockStartPt = terminalBody
    ? terminalBody.flowBounds.yPt + terminalBody.flowBounds.heightPt
    : bodyDomain.logicalBounds.yPt;
  const pageFootnoteTopPt = page.layers.notes
    .filter((node) => node.kind === 'note' && node.source.story === 'footnote')
    .reduce(
      (top, note) => Math.min(top, note.flowBounds.yPt),
      bodyDomain.logicalBounds.yPt + bodyDomain.logicalBounds.heightPt,
    );
  const blockEndPt = Math.min(
    bodyDomain.logicalBounds.yPt + bodyDomain.logicalBounds.heightPt,
    pageFootnoteTopPt,
  );
  const id = `endnotes:page:${page.pageIndex}`;
  try {
    const notes = session.layoutNotes({
      kind: 'endnote',
      referenceIds: Object.freeze([...referenceIds]),
      pageIndex: page.pageIndex,
      section: endnoteRegion?.section ?? page.section,
      container: {
        id,
        kind: 'endnote',
        bounds: {
          xPt: bodyDomain.logicalBounds.xPt,
          yPt: blockStartPt,
          widthPt: bodyDomain.logicalBounds.widthPt,
          heightPt: Math.max(0, blockEndPt - blockStartPt),
        },
      },
      firstOnPage: true,
    });
    if (notes.length === 0) return layout;
    const advancePt = notes.reduce((sum, note) => sum + note.advancePt, 0);
    const endnoteLogicalBounds = Object.freeze({
      xPt: bodyDomain.logicalBounds.xPt,
      yPt: blockStartPt,
      widthPt: bodyDomain.logicalBounds.widthPt,
      heightPt: advancePt,
    });
    const endnoteDomain = Object.freeze({
      id,
      kind: 'endnote' as const,
      ...(endnoteRegion ? { sectionRegionId: endnoteRegion.id } : {}),
      logicalBounds: endnoteLogicalBounds,
      physicalBounds: endnoteRegion
        ? Object.freeze(transformRect(
            endnoteRegion.coordinateSpace.logicalToPhysical,
            endnoteLogicalBounds,
          ))
        : endnoteLogicalBounds,
    });
    const entries: PageLayerNode[] = page.layers.roots.map((entry) => entry);
    let insertionIndex = -1;
    for (let index = 0; index < entries.length; index += 1) {
      if (entries[index]!.layer === 'body') insertionIndex = index;
    }
    insertionIndex += 1;
    entries.splice(insertionIndex, 0, ...notes.map((node): PageLayerNode => ({
      layer: 'notes',
      node,
      coordinateSpace: 'section-logical',
    })));
    const bodyReadingIds = new Set(page.layers.body.map((node) => node.id));
    let readingIndex = -1;
    for (let index = 0; index < page.readingOrder.length; index += 1) {
      if (bodyReadingIds.has(page.readingOrder[index]!)) readingIndex = index;
    }
    readingIndex += 1;
    const readingOrder = [...page.readingOrder];
    readingOrder.splice(readingIndex, 0, ...notes.map((note) => note.id));
    const pages = [...layout.pages];
    pages[pageIndex] = Object.freeze({
      ...page,
      flowDomains: Object.freeze([...page.flowDomains, endnoteDomain]),
      layers: createPageLayers(entries),
      readingOrder: Object.freeze(readingOrder),
    });
    return Object.freeze({ ...layout, pages: Object.freeze(pages) });
  } catch (error) {
    if (!(error instanceof NoteCapacityExceededError)
      || error.kind !== 'endnote'
      || error.pageIndex !== page.pageIndex
      || error.containerId !== id) {
      throw error;
    }
    return Object.freeze({
      ...layout,
      diagnostics: Object.freeze([...layout.diagnostics, Object.freeze({
        code: 'UNSUPPORTED_FEATURE' as const,
        severity: 'error' as const,
        source: Object.freeze({
          story: 'endnote' as const,
          storyInstance: referenceIds[0]!,
          path: Object.freeze([]),
        }),
        message: `Document-end notes do not fit the retained terminal flow region: ${
          error instanceof Error ? error.message : String(error)
        }`,
      })]),
    });
  }
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

function pageAnchorDestinationPlan(layout: DocumentLayout) {
  const destinations = new Map<string, Readonly<{
    occurrenceId: string;
    paragraphSource: SourceRef;
    pageIndex: number;
    flowDomainId: string;
  }>>();
  for (const page of layout.pages) {
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph') continue;
      for (const drawing of node.drawings) {
        const anchor = drawing.anchorLayer;
        if (!anchor
          || anchor.horizontalOwnership !== 'page'
          || anchor.verticalOwnership !== 'page') continue;
        const occurrenceId = anchor.acquisitionOccurrenceId ?? anchor.occurrenceId;
        destinations.set(occurrenceId, Object.freeze({
          occurrenceId,
          paragraphSource: node.source,
          pageIndex: page.pageIndex,
          flowDomainId: node.flowDomainId,
        }));
      }
    }
  }
  return destinations;
}

function anchorPlanIdentity(plan: ReadonlyMap<string, unknown>): string {
  return JSON.stringify([...plan].sort(([left], [right]) => left.localeCompare(right)));
}

function* paginateBodyWithAnchorConvergenceSteps(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
  balancePlan: BodyBalancePlan,
  observer?: BodyPaginationPassObserver,
): PaginationSteps<BodyPaginationPassResult> {
  const hasPageOwnedAnchors = input.sequence.some((entry) => (
    entry.kind === 'body-block'
    && entry.block.kind === 'paragraph'
    && (entry.block.pageOwnedAnchorOccurrenceIds?.length ?? 0) > 0
  ));
  if (!hasPageOwnedAnchors) {
    return yield* paginateBodyPassSteps(
      input, services, options, reserves, null, balancePlan, observer,
    );
  }
  try {
    return (yield* convergeExactStateSteps<Readonly<{
      pass: BodyPaginationPassResult;
      plan: ReturnType<typeof pageAnchorDestinationPlan>;
    }>, number>({
      step: function* anchorPass(previous) {
        const pass = yield* paginateBodyPassSteps(
          input,
          services,
          options,
          reserves,
          previous?.plan ?? null,
          balancePlan,
          previous === undefined ? observer : undefined,
        );
        return Object.freeze({
          pass,
          plan: pageAnchorDestinationPlan(pass.layout),
        });
      },
      stateOf: (value) => anchorPlanIdentity(value.plan),
      limit: 16,
    })).value.pass;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        error.reason === 'cycle'
          ? 'Page-anchor destination acquisition repeated an exact-state cycle'
          : 'Page-anchor destination acquisition reached the operational pass limit 16',
      );
    }
    throw error;
  }
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
  observer?: BodyPaginationPassObserver,
): PaginationSteps<BodyPaginationPassResult> {
  let plan: BodyBalancePlan = new Map();
  let pass = yield* paginateBodyWithAnchorConvergenceSteps(
    input,
    services,
    options,
    reserves,
    plan,
    observer,
  );
  if (pass.terminalDiagnostic !== null) return pass;
  for (const boundary of continuousBalanceBoundaries(input)) {
    const baseline = sharedContinuousBoundaryPage(
      pass.layout,
      boundary.outgoingSectionOccurrenceId,
      boundary.incomingSectionOccurrenceId,
    );
    if (baseline === null || baseline.outgoing.flowDomainIds.length < 2) continue;
    const pageIndex = baseline.page.pageIndex;
    const targetPt = exactRetainedColumnBalanceTarget(
      input,
      pass.allocations,
      pass.footnoteReserveByPage,
      baseline.page,
      baseline.outgoing,
    );
    const nextPlan = new Map(plan);
    nextPlan.set(boundary.outgoingSectionOccurrenceId, Object.freeze({
      pageIndex,
      targetPt,
    }));
    plan = nextPlan;
    pass = yield* paginateBodyWithAnchorConvergenceSteps(
      input,
      services,
      options,
      reserves,
      plan,
    );
    if (pass.terminalDiagnostic !== null) return pass;
  }
  return pass;
}

/** Compose one retained pass through the same layout-to-paint boundary used by
 * the authoritative result. A progressive snapshot omits document-end notes:
 * their physical owner is unknowable until the terminal page exists. */
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
  const retainedEndnoteIds = includeDocumentEndnotes
    ? new Set(bodyComposed.pages.flatMap((page) =>
        page.layers.body.flatMap((node) => (
          node.kind === 'paragraph' || node.kind === 'table'
            ? endnoteIdsInRetainedSlice(node)
            : []
        ))))
    : new Set<string>();
  const authoredEndnoteIds = (input.endnoteIds ?? [])
    .filter((id) => retainedEndnoteIds.has(id));
  const endnoteStories = composeDocumentEndnotes(
    composed,
    pass.session,
    authoredEndnoteIds,
  );
  const withEndnotes = authoredEndnoteIds.length > 0
    && noteLayoutSettings.endnotePosition !== 'docEnd'
    ? appendUnsupportedNotePositionDiagnostic(
        endnoteStories,
        'endnote',
        noteLayoutSettings.endnotePosition,
        'docEnd',
      )
    : endnoteStories;
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
  let nextPublicationPages = 1;
  let publicationFailed = false;
  const passObserver: BodyPaginationPassObserver | undefined = observer
    ? {
        shouldPublish: (committedPages) => (
          !publicationFailed && committedPages >= nextPublicationPages
        ),
        publish: (pass, processedEntries) => {
          try {
            const composed = composeBodyPaginationResult(pass, input, owners, options, false);
            // The newest committed page still owns the live transition edge:
            // section-region and page-final composition can change when the
            // following page becomes committed. Keep one page as the checkpoint
            // guard and publish only the stable prefix before it.
            const publishable = Object.freeze({
              ...composed,
              pages: Object.freeze(composed.pages.slice(0, -1)),
            }) as DocumentLayout;
            if (publishable.pages.length > 0) {
              observer.onPages(publishable, processedEntries);
            }
            nextPublicationPages = Math.max(
              pass.layout.pages.length + 1,
              pass.layout.pages.length * 2,
            );
          } catch {
            // A provisional snapshot is best-effort. The same live pagination
            // session remains authoritative and must be allowed to finish.
            publicationFailed = true;
          }
        },
      }
    : undefined;
  const seed = yield* paginateBodyWithColumnBalancingSteps(
    input, services, options, [], passObserver,
  );
  const converged = (yield* convergeHeaderFooterReserveSteps<
    BodyPaginationPassResult,
    number
  >({
    seed,
    measure: (pass) => headerFooterReserves(pass, owners),
    repaginate: function* reserveRepagination(reserves, current) {
      const contexts = paginationFieldPageContexts(current.layout);
      const iterationServices = createFieldAcquisitionServicesView(services, {
        totalPages: current.layout.pages.length,
        resolveDestinationPage: (pageIndex) => contexts[pageIndex],
      });
      return yield* paginateBodyWithColumnBalancingSteps(
        input,
        iterationServices,
        options,
        reserves,
      );
    },
    identity: (pass) => paginationFieldPageContexts(pass.layout),
    requiresConvergence: seed.session.hasPaginationFields,
  })).result;
  return assertAndDeepFreezeDocumentLayout(
    composeBodyPaginationResult(converged, input, owners, options, true),
  ) as DocumentLayout;
}

export function paginateBody(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
): DocumentLayout {
  return drainPagination(paginateBodySteps(input, services, options));
}
