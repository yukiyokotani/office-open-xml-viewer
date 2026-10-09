import { FLOAT_OVERLAP_EPS } from '../float-layout.js';
import { resolveFloatingTableBoxPt } from '../float-table-geometry.js';
import { wordGridFramePreservesEmptyCarrierReference } from './table-compatibility.js';
import { frameWrapExclusionMode } from '../frame-geometry.js';
import type { FramePr } from '../types.js';
import type {
  DrawingMLCollisionRegistrySnapshotPt,
  LayoutServices,
  FloatRegistryEntryPt,
  FloatRegistrySnapshotPt,
  FloatingTablePlacementLayout,
  LayoutRect,
  ParagraphLayout,
  SourceRef,
  TableLayout,
  TableLayoutInput,
  TableRowLayoutInput,
} from './types.js';
import {
  beginFloatingTablePlacementTransaction,
  floatingTableRegistryDelta,
  resolveFloatingTablePlacementInTransaction,
} from './floating-table-transaction.js';
import {
  floatRegistryParticipant,
  resolveBlockFlowAdmission,
  resolvePageAnchoredTableDeferral,
} from './floats.js';
import { ExactConvergenceError, convergeExactState } from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutSession,
  BodyTableContinuationCursor,
  OwnerSegmentEntry,
} from './body-layout-kernel.js';
import { retainedTableRecord } from './acquisition-state.js';
import { type RetainedTableAcquisition } from './table-acquisition.js';
import { combineAdjacentTableLayoutInputs } from './adjacent-table-layout-input.js';
import { layoutTable as layoutRetainedTableInput } from './table.js';
import {
  acceptTableAnchorReferenceRefinement,
  startTableFragmentCursor,
  takeTableFragment,
  type PageDependentTableBlockRequest,
  type TableFragmentContext,
} from './table-pagination.js';
import type { BodyAcquisitionState, PhysicalAnchorFrame } from './acquisition-context.js';
import { physicalToLogicalMatrix, sectionWritingMode, transformRect } from './coordinate-space.js';
import {
  bodyOccurrenceKey,
  bodyRootFloatingTablePlacementKey,
  tableFragmentStartKey,
} from './source-key.js';
import {
  finishOwnerHostLayout,
  ownerHostFrameBox,
  ownerHostOccupiedBox,
  ownerSegmentAt,
  ownerSegmentHeaderPrefix,
  ownerSegmentInput,
  tableOwnerSegments,
  type OwnerHostFrame,
  type TableOwnerSegment,
} from './table-owner-runs.js';
import { solveExactTranslation, type StoryPageFrames } from './story-page-frames.js';
import type { LayoutFlowBlock, LayoutSourceStore } from './layout-source-store.js';
import type { TableLayoutSource } from './table-source-acquisition.js';

export interface BodyTableMeasurementContext {
  readonly dependencies: { readonly source: LayoutSourceStore };
  readonly services: LayoutServices;
  readonly state: BodyAcquisitionState;
  readonly sessionState: {
    location: BodyAcquisitionLocation;
    floatRegistry: FloatRegistrySnapshotPt;
    drawingCollisionRegistry: DrawingMLCollisionRegistrySnapshotPt;
  };
  readonly effectiveTablePositioning: LayoutSourceStore['acquisition']['effectiveTablePositioning'];
}

export interface BodyTableMeasurementOperations {
  readonly setBodyAcquisitionLocation: (
    services: LayoutServices,
    state: BodyAcquisitionState,
    sessionState: BodyTableMeasurementContext['sessionState'],
    next: BodyAcquisitionLocation,
  ) => void;
  readonly sourceElement: (store: LayoutSourceStore, ref: SourceRef) => LayoutFlowBlock;
  readonly computeTablePtLayout: (
    state: BodyAcquisitionState,
    table: TableLayoutSource,
    contentWPt: number,
    sourceIndex: number,
  ) => void;
  readonly computeAdjacentTablePtLayouts: (
    state: BodyAcquisitionState,
    members: readonly Readonly<{ table: TableLayoutSource; sourceIndex: number }>[],
    contentWPt: number,
  ) => readonly RetainedTableAcquisition[];
  readonly ordinaryAcquisitionInputForAdjacentGroup: (
    group: ReturnType<typeof combineAdjacentTableLayoutInputs>,
  ) => TableLayoutInput;
  readonly reacquireBodyTableBlock: (
    state: BodyAcquisitionState,
    store: LayoutSourceStore,
    request: PageDependentTableBlockRequest,
  ) => ParagraphLayout | TableLayout;
}

/** The state that acquires one top-level body table. A table without
 * authored or effective positioning on a vertical page is placed upright in
 * the physical page (identity paint root), so it owns its cells' physical
 * frame: the table's own decision keeps the section state, while its cell
 * content is acquired in that upright frame (`createCellState`). Every other
 * table returns the section state unchanged. Measurement and look-ahead use
 * this one decision, so their retained cell layouts agree. */
export function bodyTableAcquisitionState(
  state: BodyAcquisitionState,
  table: TableLayoutSource,
  effectiveTablePositioning: BodyTableMeasurementContext['effectiveTablePositioning'],
): BodyAcquisitionState {
  const upright = state.verticalPhys !== undefined
    && !state.acquisitionInputs.tableFormatInput(table).positioning
    && !effectiveTablePositioning(table);
  return upright ? { ...state, uprightPhysicalTable: true } : state;
}

/**
 * Physical origin of an upright table (or owner segment) occupying the
 * logical block band [cursor y, cursor y + grid width] of a vertical section
 * (§17.6.20). The Transitional clockwise frame maps it to physical x = page
 * width − logical y − grid width and starts it at the column top (physical
 * y = logical x). A native BtoT frame maps logical y to physical x unchanged
 * and its column starts at the bottom (physical y = page height − logical x),
 * so the block ends there. That native placement is generic library policy
 * transformed through the section's own frame. No comparison with a
 * Word-produced reference has established this placement.
 */
function uprightBlockPhysicalOrigin(
  physical: PhysicalAnchorFrame,
  cursorPt: Readonly<{ xPt: number; yPt: number }>,
  gridWidthPt: number,
  physicalHeightPt: number,
): Readonly<{ xPt: number; yPt: number }> {
  return physical.nativeSectionFlow == null
    ? { xPt: physical.physicalPageWidthPt - cursorPt.yPt - gridWidthPt, yPt: cursorPt.xPt }
    : { xPt: cursorPt.yPt, yPt: physical.pageHeight - cursorPt.xPt - physicalHeightPt };
}

/** Physical top of the current column, where a text-anchored frame's
 * vertical band starts: the logical inline start in the clockwise frame and
 * the logical inline end in a native counter-clockwise frame, as for
 * DrawingML anchors (production-body-layout.ts physicalColumnTopPt). */
function uprightColumnPhysicalTopPt(
  physical: PhysicalAnchorFrame,
  cursorPt: Readonly<{ xPt: number }>,
  inlineExtentPt: number,
): number {
  return physical.nativeSectionFlow == null
    ? cursorPt.xPt
    : physical.pageHeight - (cursorPt.xPt + inlineExtentPt);
}

/** A rectangle of the upright physical page in the section's logical frame:
 * the established clockwise inverse, or the inverse of the native section's
 * own canonical matrix (sectionWritingMode, never the token alone). */
function uprightPhysicalRectToLogical(
  physical: PhysicalAnchorFrame,
  section: BodyAcquisitionState['sectionLayout'],
  rect: LayoutRect,
): LayoutRect {
  if (physical.nativeSectionFlow == null) {
    return Object.freeze({
      xPt: rect.yPt,
      yPt: physical.physicalPageWidthPt - rect.xPt - rect.widthPt,
      widthPt: rect.heightPt,
      heightPt: rect.widthPt,
    });
  }
  return Object.freeze(transformRect(
    physicalToLogicalMatrix(sectionWritingMode(section), {
      widthPt: physical.pageWidth,
      heightPt: physical.pageHeight,
    }),
    rect,
  ));
}

/** The destination page and margin rectangles of the body page. */
export function bodyPageFrames(state: BodyAcquisitionState): StoryPageFrames {
  return {
    page: { xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: state.pageH },
    margin: {
      xPt: state.marginLeft,
      yPt: state.marginTop,
      widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
      heightPt: Math.max(0, state.pageH - state.marginTop - state.marginBottom),
    },
  };
}

/** Page placement handed to the pagination of a body-root table, through
 * which the page-placed content below its cells (positioned tables under
 * in-flow nested tables, text boxes holding them) is placed. Its paragraphs
 * are reacquired by the table's acquisition owner (`owner`: the upright
 * physical owner for an upright table, otherwise the section state). */
function nestedPagePlacement(
  context: BodyTableMeasurementContext,
  operations: BodyTableMeasurementOperations,
  frames: StoryPageFrames,
  translationPt: Readonly<{ xPt: number; yPt: number }>,
  owner: BodyAcquisitionState = context.state,
): NonNullable<TableFragmentContext['pagePlacement']> {
  const { dependencies } = context;
  return Object.freeze({
    frames: Object.freeze({ page: frames.page, margin: frames.margin }),
    translationPt: Object.freeze({ ...translationPt }),
    reacquireParagraph: (request: PageDependentTableBlockRequest) =>
      operations.reacquireBodyTableBlock(owner, dependencies.source, request),
  });
}

export function measureBodyTableEntry(
  context: BodyTableMeasurementContext,
  request: Parameters<NonNullable<BodyLayoutSession['measureTable']>>[0],
  operations: BodyTableMeasurementOperations,
): ReturnType<NonNullable<BodyLayoutSession['measureTable']>> {
  const { dependencies, services, state, sessionState, effectiveTablePositioning } = context;
  operations.setBodyAcquisitionLocation(services, state, sessionState, request.location);
  if (request.input.kind === 'adjacent-table-group') {
    return measureAdjacentTableGroup(context, request as AdjacentBodyTableRequest, operations);
  }
  const table = operations.sourceElement(dependencies.source, request.input.source);
  if (table.type !== 'table') throw new Error('Table source kind mismatch');
  const sourceIndex = request.input.source.path[0]!;
  const authoredPositioning = state.acquisitionInputs.tableFormatInput(table).positioning;
  const acquisitionState = bodyTableAcquisitionState(state, table, effectiveTablePositioning);
  const uprightPhysical = acquisitionState.uprightPhysicalTable === true;
  operations.computeTablePtLayout(acquisitionState, table, request.availableInlineExtentPt, sourceIndex);
  const retained = retainedTableRecord(state, sourceIndex).acquisition;
  if (request.cursor && request.cursor.kind !== 'table') {
    throw new Error('Ordinary table acquisition received an adjacent-group cursor');
  }
  const cursor = request.cursor?.cursor ?? startTableFragmentCursor();
  const pageHeightPt = state.pageH;
  if (authoredPositioning) {
    return measurePositionedTable(context, request as OrdinaryBodyTableRequest, operations, {
      table,
      sourceIndex,
      retained,
      cursor,
      pageHeightPt,
      authoredPositioning,
    });
  }
  const segments = tableOwnerSegments(retained.input);
  if (uprightPhysical && state.verticalPhys) {
    if (segments) {
      return measureUprightOwnerSegment(context, request, operations, {
        retained,
        segments,
        cursor,
        continuation: (tableCursor, entry) => Object.freeze({
          kind: 'table' as const,
          cursor: tableCursor,
          ...entry,
        }),
      }, acquisitionState);
    }
    if (request.cursor) {
      throw new Error('An upright physical table must remain atomic');
    }
    const physical = state.verticalPhys;
    const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
    if (
      tableWidthPt > request.availableBlockExtentPt &&
      request.availableBlockExtentPt < request.freshPageBlockExtentPt
    ) {
      return Object.freeze({
        layout: retained.layout,
        blockExtentPt: 0,
        nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
        requiresFreshFlowRegion: true,
      });
    }
    const origin = uprightBlockPhysicalOrigin(
      physical, request.location.cursorPt, tableWidthPt, retained.layout.advancePt,
    );
    const physicalLeftPt = origin.xPt;
    const physicalTopPt = origin.yPt;
    const upright = takeUprightTableFragment(context, request, operations, retained, {
      xPt: physicalLeftPt,
      yPt: physicalTopPt,
    }, acquisitionState);
    return Object.freeze({
      layout: upright,
      blockExtentPt: tableWidthPt,
      nextCursor: null,
      placement: Object.freeze({
        coordinateSpace: 'upright-physical' as const,
        xPt: physicalLeftPt + upright.flowBounds.xPt,
        yPt: physicalTopPt + upright.flowBounds.yPt,
        sectionFlowOwnership: 'host-flow' as const,
      }),
    });
  }
  if (segments) {
    return measureOwnerSegmentedTable(context, request, operations, {
      retained,
      segments,
      cursor,
      continuation: (tableCursor, entry) => Object.freeze({
        kind: 'table' as const,
        cursor: tableCursor,
        ...entry,
      }),
    });
  }
  const taken = takeInFlowTableFragment(context, request, operations, retained, cursor);
  if (taken.kind === 'retry') {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: request.cursor ?? null,
      retryAtBlockStartPt: taken.retryAtBlockStartPt,
    });
  }
  if (taken.kind === 'fresh-flow-region') {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
      requiresFreshFlowRegion: true,
    });
  }
  const { result, fragment } = taken;
  return Object.freeze({
    layout: fragment,
    ...inFlowFragmentCharge(result, fragment),
    nextCursor: result.nextCursor
      ? Object.freeze({ kind: 'table' as const, cursor: result.nextCursor })
      : null,
  });
}

/** A §17.6.20 upright physical table (or owner segment) as one atomic
 * fragment in the physical page frame, translated to `translationPt`. Its
 * page-dependent cell content is reacquired by `owner`, the same upright
 * physical acquisition owner that measured it (bodyTableAcquisitionState). */
function takeUprightTableFragment(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  retained: RetainedTableAcquisition,
  translationPt: Readonly<{ xPt: number; yPt: number }>,
  owner: BodyAcquisitionState,
) {
  const { dependencies, services, state } = context;
  const physical = state.verticalPhys;
  if (!physical) throw new Error('An upright table requires a vertical section');
  const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const physicalBandHeightPt = Math.max(
    retained.layout.advancePt,
    physical.pageHeight - physical.marginTop - physical.marginBottom,
  );
  const flowDomainId = `upright-physical-page:${request.location.pageIndex}`;
  const physicalMargin = {
    xPt: physical.marginLeft,
    yPt: physical.marginTop,
    widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
    heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
  };
  const physicalPage = { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight };
  const upright = takeTableFragment(retained, startTableFragmentCursor(), {
    availableHeightPt: physicalBandHeightPt,
    freshPageHeightPt: physicalBandHeightPt,
    placement: {
      container: {
        id: flowDomainId,
        kind: 'body',
        bounds: { xPt: 0, yPt: 0, widthPt: tableWidthPt, heightPt: physicalBandHeightPt },
      },
      cursor: { xPt: 0, yPt: 0 },
      availableBounds: {
        xPt: 0,
        yPt: 0,
        widthPt: tableWidthPt,
        heightPt: physicalBandHeightPt,
      },
    },
    services,
    compatibility: 'word',
    oversizedRowPolicy: 'atomic',
    page: {
      physicalPageIndex: request.location.pageIndex,
      displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
      occurrenceId: `${retained.input.id}:upright-page:${request.location.pageIndex}`,
    },
    floatingTableFrames: {
      page: physicalPage,
      margin: physicalMargin,
      column: { ...physicalMargin },
    },
    floatingTableRegistry: Object.freeze({
      coordinateSpace: 'upright-physical-page-points' as const,
      flowDomainId,
      entries: Object.freeze([]),
      nextParagraphId: 0,
    }),
    finalPlacementTranslationPt: translationPt,
    reacquirePageDependentBlock: (blockRequest) =>
      operations.reacquireBodyTableBlock(owner, dependencies.source, blockRequest),
    pagePlacement: nestedPagePlacement(context, operations, {
      page: physicalPage,
      margin: physicalMargin,
    }, translationPt, owner),
  });
  if (!upright.fragment || upright.nextCursor || upright.requiresFreshPage) {
    throw new Error('Upright table final-frame layout must remain atomic');
  }
  return upright.fragment;
}

type InFlowTableFragment =
  | Readonly<{ kind: 'retry'; retryAtBlockStartPt: number }>
  | Readonly<{ kind: 'fresh-flow-region' }>
  | Readonly<{
      kind: 'fragment';
      result: ReturnType<typeof takeTableFragment>;
      fragment: NonNullable<ReturnType<typeof takeTableFragment>['fragment']>;
    }>;

/** One ordinary-flow fragment of a retained table at the request cursor,
 * admitted against the committed page registry like any block. */
function takeInFlowTableFragment(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  retained: RetainedTableAcquisition,
  cursor: ReturnType<typeof startTableFragmentCursor>,
): InFlowTableFragment {
  const { dependencies, services, state, sessionState } = context;
  const pageHeightPt = state.pageH;
  const fragmentContext: TableFragmentContext = {
    availableHeightPt: request.availableBlockExtentPt,
    freshPageHeightPt: request.freshPageBlockExtentPt,
    placement: {
      container: {
        id: request.location.flowDomainId,
        kind: 'body',
        bounds: {
          xPt: 0,
          yPt: 0,
          widthPt: request.availableInlineExtentPt,
          heightPt: request.availableBlockExtentPt,
        },
      },
      cursor: { xPt: 0, yPt: 0 },
      availableBounds: {
        xPt: 0,
        yPt: 0,
        widthPt: request.availableInlineExtentPt,
        heightPt: request.availableBlockExtentPt,
      },
    },
    services,
    compatibility: 'word',
    page: {
      physicalPageIndex: request.location.pageIndex,
      displayPageNumber: request.location.pageIndex + 1,
      occurrenceId: `${retained.input.id}:body:${request.location.pageIndex}`,
    },
    floatingTableFrames: {
      page: { xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt },
      margin: {
        xPt: state.marginLeft,
        yPt: state.marginTop,
        widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
        heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
      },
      column: request.location.availableBounds,
    },
    floatingTableRegistry: sessionState.floatRegistry,
    finalPlacementTranslationPt: {
      xPt: request.location.availableBounds.xPt,
      yPt: request.location.cursorPt.yPt,
    },
    reacquirePageDependentBlock: (request) =>
      operations.reacquireBodyTableBlock(state, dependencies.source, request),
    pagePlacement: nestedPagePlacement(context, operations, bodyPageFrames(state), {
      xPt: request.location.availableBounds.xPt,
      yPt: request.location.cursorPt.yPt,
    }),
  };
  let result = takeTableFragment(retained, cursor, fragmentContext);
  const tableInlineStartPt = request.location.availableBounds.xPt + retained.layout.flowBounds.xPt;
  const tableInlineEndPt = tableInlineStartPt + retained.layout.flowBounds.widthPt;
  const remainingTableExtentPt = result.fragment?.advancePt ?? 0;
  const retryAtBlockStartPt = resolveBlockFlowAdmission({
    inlineStartPt: tableInlineStartPt,
    inlineEndPt: tableInlineEndPt,
    flowBandStartPt: request.location.availableBounds.xPt,
    flowBandEndPt: request.location.availableBounds.xPt + request.location.availableBounds.widthPt,
    blockStartPt: request.location.cursorPt.yPt,
    blockExtentPt: remainingTableExtentPt,
    blockers: sessionState.floatRegistry.entries.map(floatRegistryParticipant),
    overlapEpsilonPt: FLOAT_OVERLAP_EPS,
  }).blockStartPt;
  if (retryAtBlockStartPt > request.location.cursorPt.yPt) {
    return Object.freeze({ kind: 'retry' as const, retryAtBlockStartPt });
  }
  if (!result.fragment || result.requiresFreshPage) {
    return Object.freeze({ kind: 'fresh-flow-region' as const });
  }
  const initial = request.unwrappedLocation;
  if (initial && initial.flowDomainId === request.location.flowDomainId
    && initial.pageIndex === request.location.pageIndex && initial.columnIndex === request.location.columnIndex
    && state.verticalPhys === undefined && !retained.input.bidiVisual
    && retained.input.source.story === 'body'
    && retained.input.rows[0]?.cells[0] !== undefined
    && wordGridFramePreservesEmptyCarrierReference(retained.input.rows[0].cells[0])
    && initial.cursorPt.yPt < request.location.cursorPt.yPt
    && cursor.rowIndex === 0 && cursor.rowFragmentIndex === 0 && cursor.cells.length === 0) {
    // Preserve ordinary blockers. Only the cell-start grid-frame admission
    // class separates an empty carrier's reference from actual table flow.
    const nominalBlockStartPt = resolveBlockFlowAdmission({
      inlineStartPt: tableInlineStartPt, inlineEndPt: tableInlineEndPt,
      flowBandStartPt: request.location.availableBounds.xPt,
      flowBandEndPt: request.location.availableBounds.xPt + request.location.availableBounds.widthPt,
      blockStartPt: initial.cursorPt.yPt, blockExtentPt: remainingTableExtentPt,
      blockers: sessionState.floatRegistry.entries.filter((entry) => (
        entry.paragraphAnchorReference !== 'unwrapped-empty-carrier'
      )).map(floatRegistryParticipant),
      overlapEpsilonPt: FLOAT_OVERLAP_EPS,
    }).blockStartPt;
    if (nominalBlockStartPt < request.location.cursorPt.yPt) {
      const adjusted = takeTableFragment(retained, cursor, {
        ...fragmentContext,
        paragraphAnchorReferenceDeltaPt: nominalBlockStartPt - request.location.cursorPt.yPt,
      });
      result = acceptTableAnchorReferenceRefinement(result, adjusted);
    }
  }
  return Object.freeze({ kind: 'fragment' as const, result, fragment: result.fragment! });
}

function inFlowFragmentCharge(
  result: ReturnType<typeof takeTableFragment>,
  fragment: NonNullable<ReturnType<typeof takeTableFragment>['fragment']>,
) {
  return {
    blockExtentPt: fragment.advancePt,
    ...(fragment.unpaintedOverflowPt !== undefined
      ? { unpaintedOverflowPt: fragment.unpaintedOverflowPt }
      : {}),
    ...(result.floatingTableRegistryDelta
      ? {
          flowRegistryDelta: Object.freeze({
            floats: result.floatingTableRegistryDelta,
          }),
        }
      : {}),
  };
}

type BodyTableRequest = Parameters<NonNullable<BodyLayoutSession['measureTable']>>[0];
type TableMeasureResult = ReturnType<NonNullable<BodyLayoutSession['measureTable']>>;
type AdjacentBodyTableRequest = BodyTableRequest & {
  input: Extract<BodyTableRequest['input'], { kind: 'adjacent-table-group' }>;
};
type OrdinaryBodyTableRequest = BodyTableRequest & {
  input: Exclude<BodyTableRequest['input'], { kind: 'adjacent-table-group' }>;
};
interface PositionedTableFrame {
  readonly table: TableLayoutSource;
  readonly sourceIndex: number;
  readonly retained: RetainedTableAcquisition;
  readonly cursor: ReturnType<typeof startTableFragmentCursor>;
  readonly pageHeightPt: number;
  readonly authoredPositioning: NonNullable<
    ReturnType<BodyAcquisitionState['acquisitionInputs']['tableFormatInput']>['positioning']
  >;
}

function measureAdjacentTableGroup(
  context: BodyTableMeasurementContext,
  request: AdjacentBodyTableRequest,
  operations: BodyTableMeasurementOperations,
): TableMeasureResult {
  const { dependencies, services, state } = context;

  if (request.cursor && request.cursor.kind !== 'adjacent-table-group') {
    throw new Error('Adjacent table group acquisition received an ordinary table cursor');
  }
  const records = operations.computeAdjacentTablePtLayouts(state, request.input.tables.map((tableInput) => {
    const table = operations.sourceElement(dependencies.source, tableInput.source);
    if (table.type !== 'table') throw new Error('Table source kind mismatch');
    return { table, sourceIndex: tableInput.source.path[0]! };
  }), request.availableInlineExtentPt);
  const combined = combinedAdjacentGroupAcquisition(context, request, operations, records);
  const groupCursor: import('./body-layout-kernel.js').AdjacentTableGroupCursor =
    request.cursor?.cursor ??
    Object.freeze({
      tableIndex: 0,
      sourceRowIndex: 0,
    });
  const rowsBefore = adjacentGroupRowOffsets(request.input.tables)[groupCursor.tableIndex] ?? 0;
  return measureAdjacentGroupAt(context, request, operations, combined, groupCursor, rowsBefore);
}

const adjacentGroupRowOffsetCache = new WeakMap<
  AdjacentBodyTableRequest['input']['tables'],
  readonly number[]
>();

/** First combined row of each group member (and the total at the end),
 * computed once per immutable group input. */
function adjacentGroupRowOffsets(tables: AdjacentBodyTableRequest['input']['tables']): readonly number[] {
  const cached = adjacentGroupRowOffsetCache.get(tables);
  if (cached) return cached;
  const offsets = [0];
  for (const table of tables) offsets.push(offsets.at(-1)! + (table.rowCount ?? 0));
  adjacentGroupRowOffsetCache.set(tables, Object.freeze(offsets));
  return offsets;
}

function adjacentGroupPlacement(request: AdjacentBodyTableRequest) {
  return {
    container: {
      id: request.location.flowDomainId,
      kind: 'body' as const,
      bounds: {
        xPt: 0,
        yPt: 0,
        widthPt: request.availableInlineExtentPt,
        heightPt: request.freshPageBlockExtentPt,
      },
    },
    cursor: { xPt: 0, yPt: 0 },
    availableBounds: {
      xPt: 0,
      yPt: 0,
      widthPt: request.availableInlineExtentPt,
      heightPt: request.freshPageBlockExtentPt,
    },
  };
}

const adjacentGroupAcquisitions = new WeakMap<object, Readonly<{
  records: readonly RetainedTableAcquisition[];
  widthPt: number;
  heightPt: number;
  combined: RetainedTableAcquisition;
}>>();

/** The §17.4.37 group's combined acquisition. Reused while every member
 * acquisition and the frame are unchanged, so a group paged one owner segment
 * per request is combined and laid out once, not once per segment. */
function combinedAdjacentGroupAcquisition(
  context: BodyTableMeasurementContext,
  request: AdjacentBodyTableRequest,
  operations: BodyTableMeasurementOperations,
  records: readonly RetainedTableAcquisition[],
): RetainedTableAcquisition {
  const { services } = context;
  const cached = adjacentGroupAcquisitions.get(request.input);
  if (cached
    && cached.widthPt === request.availableInlineExtentPt
    && cached.heightPt === request.freshPageBlockExtentPt
    && cached.records.length === records.length
    && cached.records.every((record, index) => record === records[index])) {
    return cached.combined;
  }
  const combinedInput = operations.ordinaryAcquisitionInputForAdjacentGroup(
    combineAdjacentTableLayoutInputs(
      request.input.logicalSequenceId,
      records.map((record) => record.input),
    ),
  );
  const combinedLayout = layoutRetainedTableInput(
    combinedInput, adjacentGroupPlacement(request), services,
  ).layout;
  const nestedById: Record<string, RetainedTableAcquisition> = {};
  records.forEach((record) =>
    Object.entries(record.nestedById).forEach(([id, nested]) => {
      if (nestedById[id] && nestedById[id] !== nested) {
        throw new Error(`Adjacent table group has duplicate nested table id: ${id}`);
      }
      nestedById[id] = nested;
    }),
  );
  const combined: RetainedTableAcquisition = Object.freeze({
    input: combinedInput,
    layout: combinedLayout,
    nestedById: Object.freeze(nestedById),
    floatingTables: Object.freeze(records.flatMap((record) => record.floatingTables)),
  });
  adjacentGroupAcquisitions.set(request.input, Object.freeze({
    records: Object.freeze([...records]),
    widthPt: request.availableInlineExtentPt,
    heightPt: request.freshPageBlockExtentPt,
    combined,
  }));
  return combined;
}

function measureAdjacentGroupAt(
  context: BodyTableMeasurementContext,
  request: AdjacentBodyTableRequest,
  operations: BodyTableMeasurementOperations,
  combined: RetainedTableAcquisition,
  groupCursor: import('./body-layout-kernel.js').AdjacentTableGroupCursor,
  rowsBefore: number,
): TableMeasureResult {
  const { services } = context;
  const globalRowIndex = rowsBefore + groupCursor.sourceRowIndex;
  const cursor =
    groupCursor.tableCursor ??
    Object.freeze({
      ...startTableFragmentCursor(),
      rowIndex: globalRowIndex,
    });
  if (cursor.rowIndex !== globalRowIndex) {
    throw new Error('Adjacent-table group and table-fragment cursors disagree');
  }
  // The §17.4.37 logical table is the owner-run domain, so a Word-saved split
  // into adjacent one-row tables keeps one host per differing carrier.
  const segments = tableOwnerSegments(combined.input);
  if (segments) {
    return measureOwnerSegmentedTable(context, request, operations, {
      retained: combined,
      segments,
      cursor,
      continuation: (tableCursor, entry) => {
        const groupAt = adjacentGroupCursorAt(request.input.tables, tableCursor);
        if (!groupAt) throw new Error('Adjacent-table owner segment cursor exceeds its group');
        return Object.freeze({ kind: 'adjacent-table-group' as const, cursor: groupAt, ...entry });
      },
    });
  }
  const result = takeTableFragment(combined, cursor, {
    availableHeightPt: request.availableBlockExtentPt,
    freshPageHeightPt: request.freshPageBlockExtentPt,
    placement: adjacentGroupPlacement(request),
    services,
    compatibility: 'word',
    page: {
      physicalPageIndex: request.location.pageIndex,
      displayPageNumber: request.location.pageIndex + 1,
      occurrenceId: `${combined.input.id}:body:${request.location.pageIndex}`,
    },
    // Placed at the cursor like an in-flow table.
    pagePlacement: nestedPagePlacement(context, operations, bodyPageFrames(context.state), {
      xPt: request.location.availableBounds.xPt,
      yPt: request.location.cursorPt.yPt,
    }),
  });
  if (!result.fragment || result.requiresFreshPage) {
    return Object.freeze({
      layout: combined.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'adjacent-table-group' as const,
        cursor: groupCursor,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  const nextGroupCursor = result.nextCursor
    ? adjacentGroupCursorAt(request.input.tables, result.nextCursor)
    : null;
  return Object.freeze({
    layout: result.fragment,
    blockExtentPt: result.fragment.advancePt,
    ...(result.fragment.unpaintedOverflowPt !== undefined
      ? { unpaintedOverflowPt: result.fragment.unpaintedOverflowPt }
      : {}),
    nextCursor: nextGroupCursor
      ? Object.freeze({ kind: 'adjacent-table-group' as const, cursor: nextGroupCursor })
      : null,
    ...(result.floatingTableRegistryDelta
      ? {
          flowRegistryDelta: Object.freeze({
            floats: result.floatingTableRegistryDelta,
          }),
        }
      : {}),
  });
}

function adjacentGroupCursorAt(
  tables: AdjacentBodyTableRequest['input']['tables'],
  tableCursor: ReturnType<typeof startTableFragmentCursor>,
): import('./body-layout-kernel.js').AdjacentTableGroupCursor | null {
  // The first member whose row range ends after the cursor row (members with
  // no rows are skipped), found by binary search over the cached offsets.
  const offsets = adjacentGroupRowOffsets(tables);
  let low = 0;
  let high = tables.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (offsets[middle + 1]! > tableCursor.rowIndex) high = middle;
    else low = middle + 1;
  }
  if (low >= tables.length) return null;
  return Object.freeze({
    tableIndex: low,
    sourceRowIndex: tableCursor.rowIndex - offsets[low]!,
    tableCursor,
  });
}

function measurePositionedTable(
  context: BodyTableMeasurementContext,
  request: OrdinaryBodyTableRequest,
  operations: BodyTableMeasurementOperations,
  frame: PositionedTableFrame,
): TableMeasureResult {
  const { dependencies, services, state, sessionState } = context;
  const { table, sourceIndex, retained, cursor, pageHeightPt, authoredPositioning } = frame;

  const positioning =
    request.cursor?.kind === 'table' && request.cursor.floatingContinuationFrame === 'fresh-text'
      ? Object.freeze({ ...authoredPositioning, vertAnchor: 'text', yPt: 0, yAlign: undefined })
      : authoredPositioning;
  const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const frames = Object.freeze({
    page: Object.freeze({ xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt }),
    margin: Object.freeze({
      xPt: state.marginLeft,
      yPt: state.marginTop,
      widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
      heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
    }),
    text: Object.freeze({
      xPt: request.location.cursorPt.xPt,
      yPt: request.location.cursorPt.yPt,
      widthPt: request.availableInlineExtentPt,
      heightPt: retained.layout.advancePt,
    }),
  });
  const raw = resolveFloatingTableBoxPt(
    positioning,
    frames,
    tableWidthPt,
    retained.layout.advancePt,
  );
  const ownPrescanOccurrenceId = bodyRootFloatingTablePlacementKey(
    request.input.source,
    request.location.pageIndex,
    cursor.rowIndex,
    cursor.rowFragmentIndex,
  );
  // The advance registration is for text before this table. Its
  // nested contents must acquire against other floats, not against
  // the table that owns them.
  const hasOwnPrescan =
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    sessionState.floatRegistry.entries.some(
      (entry) => entry.occurrenceId === ownPrescanOccurrenceId,
    );
  const nestedAcquisitionRegistry = hasOwnPrescan
    ? Object.freeze({
        ...sessionState.floatRegistry,
        entries: Object.freeze(
          sessionState.floatRegistry.entries.filter(
            (entry) => entry.occurrenceId !== ownPrescanOccurrenceId,
          ),
        ),
      })
    : sessionState.floatRegistry;
  const pageAnchoredCollision =
    request.cursor?.kind !== 'table' &&
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    resolvePageAnchoredTableDeferral({
      bounds: {
        xPt: raw.x,
        yPt: raw.y,
        widthPt: raw.w,
        heightPt: raw.h,
      },
      blockers: sessionState.floatRegistry.entries
        .filter((entry) => entry.occurrenceId !== ownPrescanOccurrenceId)
        .map(floatRegistryParticipant),
      overlapEpsilonPt: FLOAT_OVERLAP_EPS,
    }).defer;
  if (pageAnchoredCollision) {
    // `word-page-anchored-table-collision-deferral`: a fresh page
    // preserves the authored absolute anchor instead of converting
    // the colliding table to a text continuation.
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'authored' as const,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  const absoluteAnchorMustSplit =
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    retained.layout.advancePt > request.freshPageBlockExtentPt;
  const admissionBlockEndPt = absoluteAnchorMustSplit
    ? request.location.availableBounds.yPt + request.location.availableBounds.heightPt
    : positioning.vertAnchor === 'page'
      ? frames.page.yPt + frames.page.heightPt
      : positioning.vertAnchor === 'margin'
        ? frames.margin.yPt + frames.margin.heightPt
        : request.location.availableBounds.yPt + request.location.availableBounds.heightPt;
  const freshAdmissionHeightPt = absoluteAnchorMustSplit
    ? request.freshPageBlockExtentPt
    : positioning.vertAnchor === 'page'
      ? frames.page.heightPt
      : positioning.vertAnchor === 'margin'
        ? frames.margin.heightPt
        : request.freshPageBlockExtentPt;
  const transaction = convergeFloatingParentTransaction(context, request, operations, {
    table,
    retained,
    cursor,
    positioning,
    frames,
    raw,
    nestedAcquisitionRegistry,
    admissionBlockEndPt,
    freshAdmissionHeightPt,
  });
  if (transaction.kind === 'fresh-flow-region') {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'fresh-text' as const,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  const { result, fragment, resolved, nestedEntries } = transaction;
  const isFloatingContinuation =
    request.cursor?.kind === 'table' && request.cursor.floatingContinuationFrame !== undefined;
  const admittedBlockEndPt =
    request.location.availableBounds.yPt + request.location.availableBounds.heightPt;
  const hostFlowPlacements = [
    ...(fragment.resolvedFloatingTables ?? []),
    resolved.placement,
  ].filter((placement) => placement.source.positioning.vertAnchor === 'text');
  if (
    !isFloatingContinuation &&
    hostFlowPlacements.some(
      (placement) =>
        placement.exclusionBounds.yPt + placement.exclusionBounds.heightPt > admittedBlockEndPt,
    )
  ) {
    return Object.freeze({
      layout: fragment,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'fresh-text' as const,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  return Object.freeze({
    layout: fragment,
    blockExtentPt: 0,
    nextCursor: result.nextCursor
      ? Object.freeze({
          kind: 'table' as const,
          cursor: result.nextCursor,
          floatingContinuationFrame: 'fresh-text' as const,
        })
      : null,
    flowRegistryDelta: Object.freeze({
      floats: floatingTableRegistryDelta(
        sessionState.floatRegistry,
        Object.freeze([...nestedEntries, ...resolved.transaction.delta]),
        resolved.transaction.nextParagraphId,
      ),
    }),
    placement: Object.freeze({
      coordinateSpace: 'logical-body' as const,
      xPt: resolved.placement.xPt,
      yPt: resolved.placement.yPt,
      sectionFlowOwnership:
        positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin'
          ? ('page' as const)
          : ('host-flow' as const),
    }),
  });
}

type FloatingParentTransactionPass =
  | Readonly<{
      kind: 'fresh-flow-region';
      result: ReturnType<typeof takeTableFragment>;
    }>
  | Readonly<{
      kind: 'candidate';
      parentFrame: Readonly<{ xPt: number; yPt: number }>;
      result: ReturnType<typeof takeTableFragment>;
      fragment: NonNullable<ReturnType<typeof takeTableFragment>['fragment']>;
      resolved: ReturnType<typeof resolveFloatingTablePlacementInTransaction>;
      nestedEntries: readonly FloatRegistryEntryPt[];
      fingerprint: string;
    }>;
interface FloatingConvergenceFrame {
  readonly table: TableLayoutSource;
  readonly retained: RetainedTableAcquisition;
  readonly cursor: ReturnType<typeof startTableFragmentCursor>;
  readonly positioning: PositionedTableFrame['authoredPositioning'];
  readonly frames: Parameters<typeof resolveFloatingTableBoxPt>[1];
  readonly raw: ReturnType<typeof resolveFloatingTableBoxPt>;
  readonly nestedAcquisitionRegistry: FloatRegistrySnapshotPt;
  readonly admissionBlockEndPt: number;
  readonly freshAdmissionHeightPt: number;
}

function convergeFloatingParentTransaction(
  context: BodyTableMeasurementContext,
  request: OrdinaryBodyTableRequest,
  operations: BodyTableMeasurementOperations,
  frame: FloatingConvergenceFrame,
): FloatingParentTransactionPass {
  const { dependencies, services, state, sessionState } = context;
  const {
    table,
    retained,
    cursor,
    positioning,
    frames,
    raw,
    nestedAcquisitionRegistry,
    admissionBlockEndPt,
    freshAdmissionHeightPt,
  } = frame;
  let transaction: FloatingParentTransactionPass;
  try {
    transaction = convergeExactState<FloatingParentTransactionPass>({
      step: (previous) => {
        if (previous?.kind === 'fresh-flow-region') return previous;
        if (
          previous?.kind === 'candidate' &&
          previous.resolved.placement.xPt === previous.parentFrame.xPt &&
          previous.resolved.placement.yPt === previous.parentFrame.yPt
        ) {
          return previous;
        }
        const parentFrame = previous?.resolved.placement ?? {
          xPt: raw.x,
          yPt: raw.y,
        };
        const availableHeightPt = Math.max(0, admissionBlockEndPt - parentFrame.yPt);
        const result = takeTableFragment(retained, cursor, {
          availableHeightPt,
          freshPageHeightPt: freshAdmissionHeightPt,
          placement: {
            container: {
              id: `${request.location.flowDomainId}:floating-table`,
              kind: 'body',
              bounds: {
                xPt: 0,
                yPt: 0,
                widthPt: request.availableInlineExtentPt,
                heightPt: availableHeightPt,
              },
            },
            cursor: { xPt: 0, yPt: 0 },
            availableBounds: {
              xPt: 0,
              yPt: 0,
              widthPt: request.availableInlineExtentPt,
              heightPt: availableHeightPt,
            },
          },
          services,
          compatibility: 'word',
          oversizedRowPolicy: 'atomic',
          page: {
            physicalPageIndex: request.location.pageIndex,
            displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
            occurrenceId: `${retained.input.id}:fitting-outer:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
          },
          floatingTableFrames: {
            page: frames.page,
            margin: frames.margin,
            column: frames.text,
          },
          floatingTableRegistry: nestedAcquisitionRegistry,
          finalPlacementTranslationPt: parentFrame,
          reacquirePageDependentBlock: (request) =>
            operations.reacquireBodyTableBlock(state, dependencies.source, request),
          pagePlacement: nestedPagePlacement(context, operations, {
            page: frames.page,
            margin: frames.margin,
          }, parentFrame),
        });
        if (!result.fragment || result.requiresFreshPage) {
          return Object.freeze({
            kind: 'fresh-flow-region' as const,
            result,
          });
        }
        const sourcePlacement: FloatingTablePlacementLayout = Object.freeze({
          kind: 'floating-table-placement',
          occurrenceId: bodyRootFloatingTablePlacementKey(
            request.input.source,
            request.location.pageIndex,
            cursor.rowIndex,
            cursor.rowFragmentIndex,
          ),
          ownership: 'source',
          physicalPageIndex: request.location.pageIndex,
          displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
          hostCellId: request.location.flowDomainId,
          sourceBlockIndex: request.input.source.path[0]!,
          anchorBlockIndex: request.input.source.path[0]!,
          tableId: result.fragment.id,
          overlap: table.overlap === 'never' ? 'never' : 'overlap',
          positioning,
          anchorBounds: frames.text,
          child: result.fragment,
        });
        const nestedEntries = result.floatingTableRegistryDelta?.entries ?? [];
        const nestedNextParagraphId =
          result.floatingTableRegistryDelta?.nextParagraphId ??
          sessionState.floatRegistry.nextParagraphId;
        const resolved = resolveFloatingTablePlacementInTransaction(
          sourcePlacement,
          frames,
          beginFloatingTablePlacementTransaction(
            sessionState.floatRegistry.entries,
            nestedNextParagraphId,
            sessionState.floatRegistry.coordinateSpace,
            sessionState.floatRegistry.flowDomainId,
          ),
        );
        const fingerprint = JSON.stringify({
          parentFrame: {
            xPt: resolved.placement.xPt,
            yPt: resolved.placement.yPt,
          },
          fragment: result.fragment,
          nestedEntries,
          resolvedBounds: resolved.placement.bounds,
        });
        return Object.freeze({
          kind: 'candidate' as const,
          parentFrame: Object.freeze({
            xPt: parentFrame.xPt,
            yPt: parentFrame.yPt,
          }),
          result,
          fragment: result.fragment,
          resolved,
          nestedEntries,
          fingerprint,
        });
      },
      stateOf: (value) =>
        value.kind === 'fresh-flow-region' ? 'fresh-flow-region' : value.fingerprint,
      limit: 16,
    }).value;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        error.reason === 'cycle'
          ? 'Floating table parent/child transaction repeated an exact-state cycle'
          : 'Floating table parent/child transaction reached the operational pass limit 16',
      );
    }
    throw error;
  }
  return transaction;
}

// Cell-owner runs (table-owner-runs.ts) in the paginated body. The run
// boundaries are bounded there; every rule below is a labeled library
// choice, not Word evidence.

type TableCursor = ReturnType<typeof startTableFragmentCursor>;

interface OwnerSegmentedTable {
  /** The complete source table, or its §17.4.37 logical group. */
  readonly retained: RetainedTableAcquisition;
  readonly segments: readonly TableOwnerSegment[];
  /** Cursor over the complete input rows. */
  readonly cursor: TableCursor;
  readonly continuation: (
    cursor: TableCursor,
    entry: Readonly<{
      floatingContinuationFrame?: 'fresh-text';
      ownerSegmentEntry?: OwnerSegmentEntry;
    }>,
  ) => BodyTableContinuationCursor;
}

const ownerSegmentAcquisitions = new WeakMap<
  RetainedTableAcquisition,
  Map<string, RetainedTableAcquisition>
>();

/** A segment as a retained acquisition of its own, cached per source
 * acquisition. Rows keep their cell ids, so the source table's nested
 * acquisitions and nested floating-table occurrences are shared unchanged. */
function ownerSegmentAcquisition(
  retained: RetainedTableAcquisition,
  segment: TableOwnerSegment,
  leadingRows: readonly TableRowLayoutInput[],
  widthPt: number,
  services: LayoutServices,
): RetainedTableAcquisition {
  let cache = ownerSegmentAcquisitions.get(retained);
  if (!cache) {
    cache = new Map();
    ownerSegmentAcquisitions.set(retained, cache);
  }
  const key = `${segment.rowStart}:${segment.rowEnd}:${leadingRows.length}:${widthPt}`;
  const cached = cache.get(key);
  if (cached) return cached;
  const input = ownerSegmentInput(retained.input, segment, leadingRows);
  const bounds = { xPt: 0, yPt: 0, widthPt, heightPt: 1 };
  const acquisition: RetainedTableAcquisition = Object.freeze({
    input,
    layout: layoutRetainedTableInput(input, {
      container: { id: input.flowDomainId, kind: 'tableCell', bounds },
      cursor: { xPt: 0, yPt: 0 },
      availableBounds: bounds,
    }, services).layout,
    nestedById: retained.nestedById,
    floatingTables: retained.floatingTables,
  });
  cache.set(key, acquisition);
  return acquisition;
}

/**
 * Measure the owner segment at the cursor. Ordinary segments are in-flow
 * fragments of the source table; host segments are frame-placed owners.
 * A completed segment hands the paginator the next segment's start in the
 * same flow region, so reading order stays source order.
 */
function measureOwnerSegmentedTable(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  owner: OwnerSegmentedTable,
): TableMeasureResult {
  const { retained, segments, cursor } = owner;
  const segment = ownerSegmentAt(segments, cursor.rowIndex);
  const atSegmentStart = cursor.rowIndex === segment.rowStart
    && cursor.rowFragmentIndex === 0
    && cursor.cells.length === 0;
  // Library choice: a fragment continues the source table when it resumes
  // inside its segment, or when its segment's start moved to a new region.
  // The source table's leading header rows then repeat as repeated-header
  // occurrences inside that segment's owner (in-flow table or host viewport),
  // including header rows owned by an earlier segment when the run starts
  // inside the header block; they are never placed or registered separately.
  // A segment that starts where the previous one ended is not a continuation.
  const continuesSource = !atSegmentStart || request.cursor?.ownerSegmentEntry === 'fresh-region';
  const leadingRows = continuesSource ? ownerSegmentHeaderPrefix(retained.input, segment) : [];
  const shift = segment.rowStart - leadingRows.length;
  const toLocal = (source: TableCursor): TableCursor =>
    Object.freeze({ ...source, rowIndex: source.rowIndex - shift });
  const toSource = (local: TableCursor): TableCursor =>
    Object.freeze({ ...local, rowIndex: local.rowIndex + shift });
  const nextSegment = segment.rowEnd < retained.input.rows.length
    ? owner.continuation(
        Object.freeze({ rowIndex: segment.rowEnd, rowFragmentIndex: 0, cells: Object.freeze([]) }),
        { ownerSegmentEntry: 'same-region' },
      )
    : null;
  const movedEntry = atSegmentStart && segment.rowStart > 0
    ? { ownerSegmentEntry: 'fresh-region' as const }
    : {};
  if (segment.kind === 'host') {
    return measureOwnerHost(context, request, operations, owner, Object.freeze({
      segment, leadingRows, toLocal, toSource, nextSegment, movedEntry,
    }));
  }
  const local = ownerSegmentAcquisition(
    retained, segment, leadingRows, request.availableInlineExtentPt, context.services,
  );
  const taken = takeInFlowTableFragment(context, request, operations, local, toLocal(cursor));
  if (taken.kind === 'retry') {
    return Object.freeze({
      layout: local.layout,
      blockExtentPt: 0,
      nextCursor: request.cursor ?? null,
      retryAtBlockStartPt: taken.retryAtBlockStartPt,
    });
  }
  if (taken.kind === 'fresh-flow-region') {
    return Object.freeze({
      layout: local.layout,
      blockExtentPt: 0,
      nextCursor: owner.continuation(cursor, movedEntry),
      requiresFreshFlowRegion: true,
    });
  }
  const { result, fragment } = taken;
  return Object.freeze({
    layout: fragment,
    ...inFlowFragmentCharge(result, fragment),
    nextCursor: result.nextCursor
      ? owner.continuation(toSource(result.nextCursor), {})
      : nextSegment,
  });
}

/** A fresh-text continuation keeps the frame's horizontal position and
 * restarts vertically at the new region's text top, as positioned-table
 * continuations do (existing library policy, inherited unchanged). */
function freshTextFrame(framePr: FramePr): FramePr {
  const { yAlign: _yAlign, ...horizontal } = framePr;
  return Object.freeze({ ...horizontal, vAnchor: 'text', y: 0 });
}

/**
 * A host segment: one frame-placed owner of whole source rows.
 *
 * Placement (`ownerHostFrameBox`, the paginated body's page frame): P is
 * `computeFrameBox` for the governing carrier, with content width = the
 * natural grid width (an explicit `w` overrides only the frame box, never the
 * grid) and content height T = the remaining run extent. Page/margin anchors
 * therefore inherit `clampAbsBoxIntoContainer`, an implementation-defined
 * policy that ECMA-376 does not define: a run overflowing its container
 * shifts up by the overflow, and a run taller than it pins to the container
 * start, overflows, and continues by the positioned-table split below. Text
 * anchors ride the flow cursor and are not clamped. The run is laid out once
 * as its own table at a local origin (indent and alignment relative to the
 * owner) and translated by P; T, P and the viewport stay separate.
 *
 * Reservation (library choice, not proven normative): the owner advances the
 * flow by 0 and registers one frame exclusion; following blocks — including
 * later ordinary segments of the same table — are admitted by the existing
 * exclusion rules. Observed: in the two-row source the following paragraph
 * sat above the second host, which contradicts forcing it below every host
 * bottom; the evidence does not identify the cursor or debit algorithm. The
 * exclusion is `ownerHostOccupiedBox`; wrap maps through
 * `frameWrapExclusionMode` (none and notBeside → top-and-bottom).
 * Limitation: like body paragraph frames, a page/margin-anchored owner takes
 * no part in the page-start anchor prescan, so it never re-wraps text that
 * precedes it on the same page.
 * Extent consistency: a run the fragment completes is framed by the extent
 * its fragment takes at P, solved to an exact fixed point as the story-root
 * host is (production-body-layout.ts layoutStoryOwnerHost), so the frame,
 * the painted rows and the registered exclusion are one placement. A run
 * that continues is framed by its ACQUIRED remaining extent, which is also
 * the split policy's admission measure: its later portions are taken on
 * later pages.
 *
 * Splitting and ownership (existing positioned-table policy): rows are
 * admitted from P.y to the anchor container's end (page/margin) or the
 * footnote-reserved flow band (text); only a single row taller than a fresh
 * band is atomic; a continuation re-resolves P in a fresh text frame on its
 * region. The registry entry carries this occurrence's body key, is committed
 * only with the accepted occurrence, and the same occurrence is filtered from
 * the base it resolves against; nested floating tables resolve against that
 * base before the owner's own entry is appended, so the owner never displaces
 * its own members. Rejected, fresh-region and footnote retries return no
 * committed delta.
 */
function measureOwnerHost(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  owner: OwnerSegmentedTable,
  at: Readonly<{
    segment: TableOwnerSegment;
    leadingRows: readonly TableRowLayoutInput[];
    toLocal: (cursor: TableCursor) => TableCursor;
    toSource: (cursor: TableCursor) => TableCursor;
    nextSegment: BodyTableContinuationCursor | null;
    movedEntry: Readonly<{ ownerSegmentEntry?: OwnerSegmentEntry }>;
  }>,
): TableMeasureResult {
  const { dependencies, services, state, sessionState } = context;
  const { segment, leadingRows } = at;
  const carrier = segment.carrier;
  if (!carrier) throw new Error('A cell-owner host segment requires its governing carrier');
  const continuation = request.cursor?.floatingContinuationFrame === 'fresh-text';
  const framePr = continuation ? freshTextFrame(carrier.framePr) : carrier.framePr;
  const gridWidthPt = owner.retained.input.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const viewportWidthPt = framePr.w ?? gridWidthPt;
  // Alignment applies only within free viewport width; the grid is never
  // compressed (library choice).
  const local = ownerSegmentAcquisition(
    owner.retained, segment, leadingRows, Math.max(viewportWidthPt, gridWidthPt), services,
  );
  const localCursor = at.toLocal(owner.cursor);
  const cursorYPt = request.location.cursorPt.yPt;
  const anchorContext = {
    contentX: request.location.availableBounds.xPt,
    contentW: request.availableInlineExtentPt,
    pageH: state.pageH,
    pageWidth: state.pageWidth,
    marginLeft: state.marginLeft,
    marginRight: state.marginRight,
    marginTop: state.marginTop,
    marginBottom: state.marginBottom,
  };
  const remainingExtentPt = [
    ...local.layout.rows.slice(0, leadingRows.length),
    ...local.layout.rows.slice(localCursor.rowIndex),
  ].reduce((sum, row) => sum + row.heightPt, 0);
  const frameFor = (extentPt: number) => ownerHostFrameBox(
    framePr, { anchors: anchorContext, blockOriginPt: 0 }, cursorYPt, gridWidthPt, extentPt,
  );
  const startBox = frameFor(remainingExtentPt);
  const hostFlow = framePr.vAnchor !== 'page' && framePr.vAnchor !== 'margin';
  const bandEndPt = request.location.availableBounds.yPt + request.availableBlockExtentPt;
  const absoluteMustSplit = !hostFlow && remainingExtentPt > request.freshPageBlockExtentPt;
  const admissionBlockEndPt = hostFlow || absoluteMustSplit
    ? bandEndPt
    : framePr.vAnchor === 'page' ? state.pageH : state.pageH - state.marginBottom;
  const freshAdmissionHeightPt = hostFlow || absoluteMustSplit
    ? request.freshPageBlockExtentPt
    : framePr.vAnchor === 'page'
      ? state.pageH
      : Math.max(0, state.pageH - state.marginTop - state.marginBottom);
  const ownOccurrenceId = bodyOccurrenceKey(
    request.input.source,
    request.location.flowDomainId,
    tableFragmentStartKey(request.cursor),
  );
  const ownedByThisOccurrence = (entry: FloatRegistryEntryPt) =>
    entry.occurrenceId === ownOccurrenceId
    || entry.occurrenceId.startsWith(`${ownOccurrenceId}/occurrence/`);
  const base = sessionState.floatRegistry;
  const resolutionBase = base.entries.some(ownedByThisOccurrence)
    ? Object.freeze({ ...base, entries: Object.freeze(base.entries.filter((entry) => !ownedByThisOccurrence(entry))) })
    : base;
  // The run's fragment taken with the frame origin at (x, y): rows are
  // admitted from y, and the content below its cells is placed through it.
  const takeAt = (xPt: number, yPt: number) => {
    const availableHeightPt = Math.max(0, admissionBlockEndPt - yPt);
    const ownerBounds = {
      xPt: 0,
      yPt: 0,
      widthPt: Math.max(viewportWidthPt, gridWidthPt),
      heightPt: availableHeightPt,
    };
    return takeTableFragment(local, localCursor, {
      availableHeightPt,
      freshPageHeightPt: freshAdmissionHeightPt,
      placement: {
        container: { id: `${request.location.flowDomainId}:cell-owner`, kind: 'body', bounds: ownerBounds },
        cursor: { xPt: 0, yPt: 0 },
        availableBounds: ownerBounds,
      },
      services,
      compatibility: 'word',
      oversizedRowPolicy: 'atomic',
      page: {
        physicalPageIndex: request.location.pageIndex,
        displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
        occurrenceId: `${local.input.id}:owner-host:${request.location.pageIndex}:${owner.cursor.rowIndex}:${owner.cursor.rowFragmentIndex}`,
      },
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: state.pageH },
        margin: {
          xPt: state.marginLeft,
          yPt: state.marginTop,
          widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
          heightPt: Math.max(0, state.pageH - state.marginTop - state.marginBottom),
        },
        column: request.location.availableBounds,
      },
      floatingTableRegistry: resolutionBase,
      finalPlacementTranslationPt: { xPt, yPt },
      reacquirePageDependentBlock: (blockRequest) =>
        operations.reacquireBodyTableBlock(state, dependencies.source, blockRequest),
      pagePlacement: nestedPagePlacement(context, operations, bodyPageFrames(state), { xPt, yPt }),
    });
  };
  // P depends on the run's extent T (yAlign, the page clamp) and, through
  // page-placed content below its cells (a positioned child its anchor
  // paragraph wraps around), the fragment's extent on P. A run this fragment
  // completes is framed, painted and excluded by ONE placement: the exact
  // fixed point of P.y = frame(extent of the fragment taken at P.y), as the
  // story-root host (production-body-layout.ts layoutStoryOwnerHost) solves
  // it (library consistency policy, not a Word guarantee). A run that
  // continues is framed by its acquired remaining extent, as before: its
  // later portions are taken on later pages and have no placed extent here,
  // and T is not shortened to this fragment, so admission and the split are
  // unchanged. The frame x does not depend on T. A run whose extent does not
  // move with P (no page-placed content, or none that changes a row) is
  // accepted by its first evaluation, so its placement and its work are as
  // before. The limit is a resource guard only, as for the story-root host
  // and the final-frame reflows; no fixed point fails closed.
  const extentTakenBy = (result: ReturnType<typeof takeAt>) => (
    result.fragment && !result.requiresFreshPage && !result.nextCursor
      ? result.fragment.advancePt
      : remainingExtentPt
  );
  let solved: Readonly<{ box: ReturnType<typeof frameFor>; result: ReturnType<typeof takeAt> }>;
  try {
    solved = solveExactTranslation(startBox.y, (yPt) => {
      const taken = takeAt(startBox.x, yPt);
      const next = frameFor(extentTakenBy(taken));
      return { value: { box: next, result: taken }, implied: next.y };
    }, 16);
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `body cell-owner host placement did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
  const { box, result } = solved;
  const freshRegion = (): TableMeasureResult => Object.freeze({
    layout: local.layout,
    blockExtentPt: 0,
    nextCursor: owner.continuation(owner.cursor, {
      floatingContinuationFrame: 'fresh-text',
      ...at.movedEntry,
    }),
    requiresFreshFlowRegion: true,
  });
  if (!result.fragment || result.requiresFreshPage) return freshRegion();
  const occupied = ownerHostOccupiedBox(box, framePr, result.fragment.advancePt);
  if (hostFlow && !continuation && occupied.exBottom > bandEndPt) return freshRegion();

  const layout = finishOwnerHostLayout(result.fragment, framePr, gridWidthPt, 0);

  const nestedEntries = result.floatingTableRegistryDelta?.entries ?? [];
  const paragraphId = result.floatingTableRegistryDelta?.nextParagraphId ?? base.nextParagraphId;
  const exclusionMode = frameWrapExclusionMode(framePr);
  const hostEntries: FloatRegistryEntryPt[] = occupied.w <= 0 || occupied.h <= 0
    ? []
    : [Object.freeze({
        kind: 'frame' as const,
        occurrenceId: ownOccurrenceId,
        exclusionId: ownOccurrenceId,
        paragraphId,
        exclusionMode,
        bounds: Object.freeze({ xPt: occupied.x, yPt: occupied.y, widthPt: occupied.w, heightPt: occupied.h }),
        exclusionBounds: Object.freeze({
          xPt: occupied.exLeft,
          yPt: occupied.exTop,
          widthPt: occupied.exRight - occupied.exLeft,
          heightPt: occupied.exBottom - occupied.exTop,
        }),
      })];
  const entries = [...nestedEntries, ...hostEntries];
  return Object.freeze({
    layout,
    blockExtentPt: 0,
    nextCursor: result.nextCursor
      ? owner.continuation(at.toSource(result.nextCursor), { floatingContinuationFrame: 'fresh-text' })
      : at.nextSegment,
    ...(entries.length === 0 ? {} : {
      flowRegistryDelta: Object.freeze({
        floats: floatingTableRegistryDelta(base, entries, paragraphId + hostEntries.length),
      }),
    }),
    ...(hostFlow ? { relocationBlockExtentPt: Math.max(0, occupied.y + occupied.h - cursorYPt) } : {}),
    placement: Object.freeze({
      coordinateSpace: 'logical-body' as const,
      xPt: box.x + layout.flowBounds.xPt,
      yPt: box.y + layout.flowBounds.yPt,
      sectionFlowOwnership: hostFlow ? ('host-flow' as const) : ('page' as const),
    }),
  });
}

/**
 * A §17.6.20 vertical section consumes the same owner-run projection over its
 * upright physical table (library policy). Every segment stays
 * atomic, as every upright table does, so no header row ever repeats:
 * - an ordinary segment is an upright block of its own at the flow position,
 *   advancing the vertical flow by its physical grid width, at the same
 *   physical origin an unsegmented upright table takes in the section's own
 *   frame (uprightBlockPhysicalOrigin);
 * - a host resolves its carrier in the upright physical page frame — the
 *   frame this path already gives nested floating tables: physical page and
 *   margin, a text anchor's horizontal band is the physical margin box and its
 *   vertical band starts at the physical top of the current column
 *   (uprightColumnPhysicalTopPt) — and advances the flow by 0. Its exclusion
 *   enters the logical body registry through the inverse of the section's
 *   upright projection (uprightPhysicalRectToLogical: clockwise, logical
 *   inline = physical y and logical block start = physical page width −
 *   physical right edge; native BtoT, its own counter-clockwise inverse), so
 *   notBeside blocks the vertical lines crossing the host's physical span.
 * Cell content is acquired and reacquired by `acquisitionOwner`, the
 * upright physical owner of the whole table (bodyTableAcquisitionState).
 */
function measureUprightOwnerSegment(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  owner: OwnerSegmentedTable,
  acquisitionOwner: BodyAcquisitionState,
): TableMeasureResult {
  const { services, state, sessionState } = context;
  const physical = state.verticalPhys;
  if (!physical) throw new Error('An upright owner segment requires a vertical section');
  const { retained, segments, cursor } = owner;
  const segment = ownerSegmentAt(segments, cursor.rowIndex);
  if (cursor.rowIndex !== segment.rowStart || cursor.rowFragmentIndex !== 0 || cursor.cells.length !== 0) {
    throw new Error('An upright owner segment must remain atomic');
  }
  const nextCursor = segment.rowEnd < retained.input.rows.length
    ? owner.continuation(
        Object.freeze({ rowIndex: segment.rowEnd, rowFragmentIndex: 0, cells: Object.freeze([]) }),
        { ownerSegmentEntry: 'same-region' },
      )
    : null;
  const gridWidthPt = retained.input.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  if (!segment.carrier) {
    const local = ownerSegmentAcquisition(
      retained, segment, [], request.availableInlineExtentPt, services,
    );
    if (gridWidthPt > request.availableBlockExtentPt
      && request.availableBlockExtentPt < request.freshPageBlockExtentPt) {
      return Object.freeze({
        layout: local.layout,
        blockExtentPt: 0,
        nextCursor: owner.continuation(
          cursor,
          segment.rowStart > 0 ? { ownerSegmentEntry: 'fresh-region' } : {},
        ),
        requiresFreshFlowRegion: true,
      });
    }
    const origin = uprightBlockPhysicalOrigin(
      physical, request.location.cursorPt, gridWidthPt, local.layout.advancePt,
    );
    const physicalLeftPt = origin.xPt;
    const physicalTopPt = origin.yPt;
    const layout = takeUprightTableFragment(context, request, operations, local, {
      xPt: physicalLeftPt,
      yPt: physicalTopPt,
    }, acquisitionOwner);
    return Object.freeze({
      layout,
      blockExtentPt: gridWidthPt,
      nextCursor,
      placement: Object.freeze({
        coordinateSpace: 'upright-physical' as const,
        xPt: physicalLeftPt + layout.flowBounds.xPt,
        yPt: physicalTopPt + layout.flowBounds.yPt,
        sectionFlowOwnership: 'host-flow' as const,
      }),
    });
  }
  const framePr = segment.carrier.framePr;
  const local = ownerSegmentAcquisition(
    retained, segment, [], Math.max(framePr.w ?? gridWidthPt, gridWidthPt), services,
  );
  const frame: OwnerHostFrame = {
    anchors: {
      contentX: physical.marginLeft,
      contentW: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
      pageWidth: physical.pageWidth,
      marginLeft: physical.marginLeft,
      marginRight: physical.marginRight,
      pageH: physical.pageHeight,
      marginTop: physical.marginTop,
      marginBottom: physical.marginBottom,
    },
    blockOriginPt: 0,
  };
  const columnTopPt = uprightColumnPhysicalTopPt(
    physical, request.location.cursorPt, request.availableInlineExtentPt,
  );
  const frameFor = (extentPt: number) =>
    ownerHostFrameBox(framePr, frame, columnTopPt, gridWidthPt, extentPt);
  // The frame depends on the host extent (yAlign, the page clamp) and, through
  // page-placed content below its cells (a positioned child its anchor
  // paragraph wraps around), the extent on where the host is placed on the
  // physical page. As for the horizontal body host above and the story-root
  // host (production-body-layout.ts layoutStoryOwnerHost), the host is
  // framed, painted and excluded by ONE placement: the exact fixed point of
  // box.y = frame(extent of the atomic fragment taken at box.y), starting from
  // the frame of its acquired extent (library consistency policy, not a Word
  // guarantee). The frame x does not depend on the extent. A host whose
  // extent does not move with it is accepted by its first evaluation. The
  // segment is atomic, so there is no other region to retry: the limit is a
  // resource guard only and no fixed point fails closed, restated as a plain
  // invariant so no exact-state recovery around the page can absorb it.
  const startBox = frameFor(local.layout.advancePt);
  let solved: Readonly<{
    box: ReturnType<typeof frameFor>;
    fragment: ReturnType<typeof takeUprightTableFragment>;
  }>;
  try {
    solved = solveExactTranslation(startBox.y, (yPt) => {
      const taken = takeUprightTableFragment(
        context, request, operations, local, { xPt: startBox.x, yPt }, acquisitionOwner,
      );
      const next = frameFor(taken.advancePt);
      return { value: { box: next, fragment: taken }, implied: next.y };
    }, 16);
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `upright cell-owner host placement did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
  const { box, fragment } = solved;
  const layout = finishOwnerHostLayout(fragment, framePr, gridWidthPt, 0);
  const occupied = ownerHostOccupiedBox(box, framePr, fragment.advancePt);
  const exclusionMode = frameWrapExclusionMode(framePr);
  const ownOccurrenceId = bodyOccurrenceKey(
    request.input.source,
    request.location.flowDomainId,
    tableFragmentStartKey(request.cursor),
  );
  const toLogical = (xPt: number, yPt: number, widthPt: number, heightPt: number) =>
    uprightPhysicalRectToLogical(physical, state.sectionLayout, { xPt, yPt, widthPt, heightPt });
  const base = sessionState.floatRegistry;
  const entries: FloatRegistryEntryPt[] = occupied.w <= 0 || occupied.h <= 0
    ? []
    : [Object.freeze({
        kind: 'frame' as const,
        occurrenceId: ownOccurrenceId,
        exclusionId: ownOccurrenceId,
        paragraphId: base.nextParagraphId,
        exclusionMode,
        bounds: toLogical(occupied.x, occupied.y, occupied.w, occupied.h),
        exclusionBounds: toLogical(
          occupied.exLeft,
          occupied.exTop,
          occupied.exRight - occupied.exLeft,
          occupied.exBottom - occupied.exTop,
        ),
      })];
  return Object.freeze({
    layout,
    blockExtentPt: 0,
    nextCursor,
    ...(entries.length === 0 ? {} : {
      flowRegistryDelta: Object.freeze({
        floats: floatingTableRegistryDelta(base, entries, base.nextParagraphId + entries.length),
      }),
    }),
    placement: Object.freeze({
      coordinateSpace: 'upright-physical' as const,
      xPt: box.x + layout.flowBounds.xPt,
      yPt: box.y + layout.flowBounds.yPt,
      sectionFlowOwnership: 'host-flow' as const,
    }),
  });
}

/**
 * Keep-with-next look-ahead extent of a table projected into owner runs, or
 * `null` for an unsegmented table: a host contributes no flow advance, an
 * ordinary segment its laid-out height (same choice as measurement).
 */
export function ownerSegmentedFlowExtents(
  retained: RetainedTableAcquisition,
  widthPt: number,
  services: LayoutServices,
): Readonly<{ fullExtentPt: number; leadContentExtentPt: number }> | null {
  const segments = tableOwnerSegments(retained.input);
  if (!segments) return null;
  const extents = segments.map((segment) => {
    if (segment.kind === 'host') return { advancePt: 0, leadPt: 0 };
    const { layout } = ownerSegmentAcquisition(retained, segment, [], widthPt, services);
    return { advancePt: layout.advancePt, leadPt: layout.rows[0]?.advancePt ?? layout.advancePt };
  });
  return Object.freeze({
    fullExtentPt: extents.reduce((sum, extent) => sum + extent.advancePt, 0),
    leadContentExtentPt: extents[0]?.leadPt ?? 0,
  });
}
