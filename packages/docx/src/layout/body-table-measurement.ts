import { FLOAT_OVERLAP_EPS } from '../float-layout.js';
import { resolveFloatingTableBoxPt } from '../float-table-geometry.js';
import {
  WORD_UPRIGHT_TABLE_TABLE_CLEARANCE_PT,
  WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT,
  wordGridFramePreservesEmptyCarrierReference,
} from './table-compatibility.js';
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
import { type RetainedTableAcquisition, shareRotatedCellReacquisition } from './table-acquisition.js';
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
import {
  physicalInlineBandTopPt,
  uprightBlockPhysicalOrigin,
  uprightPhysicalRectToLogical,
} from './upright-physical-frame.js';
import {
  bodyOccurrenceKey,
  bodyRootFloatingTablePlacementKey,
  anchorOccurrenceBlockIndex,
  tableFragmentStartKey,
  uprightBlockTableKey,
} from './source-key.js';
import { footnoteIdsInRetainedSlice } from './note-reference-ownership.js';
import {
  addUprightPageDrawing,
  type UprightDrawingLimits,
  emptyUprightBlockPageIndex,
  uprightBlockBoxesMet,
  uprightPageDrawingsNear,
  uprightSegmentKey,
  type UprightBlockPageIndex,
  type UprightBlockRegistration,
  type UprightBlockSlot,
} from './upright-block-index.js';
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
  // `word-vertical-section-upright-block-table`
  const upright = state.verticalPhys !== undefined
    && !state.acquisitionInputs.tableFormatInput(table).positioning
    && !effectiveTablePositioning(table);
  return upright ? { ...state, uprightPhysicalTable: true } : state;
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

/**
 * A top-level block table of a vertical section (ECMA-376 §17.6.20), laid
 * out upright: a single table, or the tables of one logical table (§17.4.37)
 * acquired as that table (`retained`), ending at body index
 * `lastSourceIndex`. A table projected into owner segments is placed segment
 * by segment; any other is an upright block table
 * (`word-vertical-section-upright-block-table`) or, in a counter-clockwise
 * section flow, an atomic block.
 */
function measureUprightTable(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  retained: RetainedTableAcquisition,
  cursor: ReturnType<typeof startTableFragmentCursor>,
  acquisitionState: BodyAcquisitionState,
  physical: PhysicalAnchorFrame,
  lastSourceIndex: number,
): TableMeasureResult {
  const segments = tableOwnerSegments(retained.input);
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
  if (uprightBlockPaginates(physical)) {
    return measureUprightBlockTable(
      context, request, operations, uprightRunMember(physical, request.input.source, retained, lastSourceIndex),
      cursor, acquisitionState, physical,
    );
  }
  // A counter-clockwise section flow (library policy): an atomic block at
  // the column bottom that advances the flow by its physical width.
  if (request.cursor) {
    throw new Error('An upright physical table must remain atomic');
  }
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
  const upright = takeUprightTableFragment(context, request, operations, retained, origin, acquisitionState);
  return Object.freeze({
    layout: upright,
    blockExtentPt: tableWidthPt,
    nextCursor: null,
    placement: Object.freeze({
      coordinateSpace: 'upright-physical' as const,
      xPt: origin.xPt + upright.flowBounds.xPt,
      yPt: origin.yPt + upright.flowBounds.yPt,
      sectionFlowOwnership: 'host-flow' as const,
    }),
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
    // The tables of a logical table (ECMA-376 §17.4.37) are one block table
    // of a vertical section and upright like any other: acquired as that one
    // table and placed by the same path, owner segments included.
    const physical = state.verticalPhys;
    const logical = physical
      ? uprightTableAcquisitionAt(context, request, operations, request.input.source.path[0]!, request.input)
      : null;
    if (physical && logical && logical.retained.input.rows.length > 0) {
      if (request.cursor && request.cursor.kind !== 'table') {
        throw new Error('An upright logical table continues by a table cursor');
      }
      return measureUprightTable(
        context, request, operations, logical.retained,
        request.cursor?.cursor ?? startTableFragmentCursor(),
        { ...state, uprightPhysicalTable: true }, physical, logical.lastSourceIndex,
      );
    }
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
  if (cursor.rowIndex >= retained.input.rows.length) {
    // Nothing is left to place: a table with no rows (CT_Tbl allows zero w:tr,
    // EG_ContentRowContent is minOccurs="0"). takeTableFragment reports this as
    // "no fragment", which every acquisition path below would read as "does not
    // fit here" and request a fresh region, and the same answer repeats on
    // every page.
    return Object.freeze({ layout: retained.layout, blockExtentPt: 0, nextCursor: null });
  }
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
    return measureUprightTable(
      context, request, operations, retained, cursor, acquisitionState, state.verticalPhys, sourceIndex,
    );
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

/** Whether an upright block of this section frame follows
 * `word-vertical-section-upright-block-table` (the vertical-rl and
 * vertical-lr frames: tbRl, tbLrV and the tokens sharing them). A
 * counter-clockwise section flow (sideways-lr) keeps the library policy of an
 * atomic block at the column bottom that advances the flow by its physical
 * width and registers no exclusion. */
export function uprightBlockPaginates(physical: PhysicalAnchorFrame): boolean {
  return physical.writingMode === 'vertical-rl' || physical.writingMode === 'vertical-lr';
}

/** The acquisition with every row kept whole: an upright table breaks only
 * between rows (`word-vertical-section-upright-block-table`), as §17.4.6
 * cantSplit rows do. */
const wholeRowAcquisitions = new WeakMap<RetainedTableAcquisition, RetainedTableAcquisition>();
function wholeRowAcquisition(retained: RetainedTableAcquisition): RetainedTableAcquisition {
  let whole = wholeRowAcquisitions.get(retained);
  if (!whole) {
    if (retained.input.rows.every((row) => row.cantSplit === true)) {
      wholeRowAcquisitions.set(retained, retained);
      return retained;
    }
    whole = Object.freeze({
      ...retained,
      input: Object.freeze({
        ...retained.input,
        rows: Object.freeze(retained.input.rows.map((row) => Object.freeze({ ...row, cantSplit: true }))),
      }),
    });
    wholeRowAcquisitions.set(retained, whole);
  }
  return whole;
}

/** The translation an upright block's fragment is laid out with when its
 * grid's near edge is placed at `originPt`: under
 * `word-vertical-section-upright-block-table` the table's indent and
 * alignment (§17.4.50, §17.4.28) do not move it, so its own inline offset is
 * dropped, and what its cells place on the page, such as floating tables, is
 * resolved without it too. */
function uprightBlockTranslationPt(
  originPt: Readonly<{ xPt: number; yPt: number }>,
  table: TableLayout,
): Readonly<{ xPt: number; yPt: number }> {
  return { xPt: originPt.xPt - table.flowBounds.xPt, yPt: originPt.yPt };
}

/** How far the painted border rules of an upright block reach past its grid
 * on its near and far block edges. */
type UprightBlockEdges = Readonly<{ nearPt: number; farPt: number }>;

/** {@link UprightBlockEdges} of a table: the outermost edges of its vertical
 * rules against the grid of physical width `physicalWidthPt`. A rule centred
 * on a grid edge reaches half its width past it; with cell spacing the outer
 * cell rules can reach further than the table's own. The near edge is the
 * physical right in the clockwise frame and the physical left in the
 * vertical-lr frame. Every fragment of a table uses the whole table's edges,
 * so the scan of its rules is kept with the retained layout. */
const uprightBlockEdgeCache = new WeakMap<TableLayout, Readonly<{ leftPt: number; rightPt: number }>>();
function uprightBlockEdgeOverhangsPt(
  table: TableLayout,
  physicalWidthPt: number,
  physical: PhysicalAnchorFrame,
): UprightBlockEdges {
  let sides = uprightBlockEdgeCache.get(table);
  if (!sides) {
    const gridLeftPt = table.flowBounds.xPt;
    const gridRightPt = gridLeftPt + physicalWidthPt;
    let leftPt = 0;
    let rightPt = 0;
    for (const border of table.borders) {
      if (border.from.xPt !== border.to.xPt) continue;
      leftPt = Math.max(leftPt, gridLeftPt - (border.from.xPt - border.widthPt / 2));
      rightPt = Math.max(rightPt, border.from.xPt + border.widthPt / 2 - gridRightPt);
    }
    sides = Object.freeze({ leftPt, rightPt });
    uprightBlockEdgeCache.set(table, sides);
  }
  return physical.writingMode === 'vertical-lr'
    ? { nearPt: sides.leftPt, farPt: sides.rightPt } : { nearPt: sides.rightPt, farPt: sides.leftPt };
}

/** The ends of the page a drawing moved for an upright table stays before
 * (uprightDrawingMovePt), in the section-logical page frame. */
export function uprightDrawingLimits(state: Pick<BodyAcquisitionState, 'pageWidth' | 'pageH'>): UprightDrawingLimits {
  return {
    inlineEndPt: state.pageWidth, blockEndPt: state.pageH, tableClearancePt: WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT,
  };
}

function uprightBlockIndexOf(state: BodyAcquisitionState): UprightBlockPageIndex {
  return state.uprightBlockIndex ?? (state.uprightBlockIndex = emptyUprightBlockPageIndex());
}

/** Physical origin of the grid of an upright block whose box is `bounds`:
 * the near rule's outer edge is the box's block start, so the grid starts
 * what that rule overhangs further on. */
function uprightBlockGridOrigin(
  physical: PhysicalAnchorFrame,
  bounds: LayoutRect,
  gridExtentPt: number,
  edges: UprightBlockEdges,
): Readonly<{ xPt: number; yPt: number }> {
  return uprightBlockPhysicalOrigin(
    physical, { xPt: bounds.xPt, yPt: bounds.yPt + edges.nearPt }, gridExtentPt, bounds.widthPt,
  );
}

/** The text exclusion of an upright block whose box is `bounds`: line strips
 * clear it only past the text clearance at its far edge. */
function uprightBlockExclusion(
  source: SourceRef,
  fragmentStartKey: string,
  bounds: LayoutRect,
): UprightBlockRegistration['exclusions'][number] {
  return Object.freeze({
    id: `${uprightBlockTableKey(source)}:${fragmentStartKey}`,
    bounds: Object.freeze({ ...bounds, heightPt: bounds.heightPt + WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT }),
  });
}

/** Body order of the fragment of `request`: its top-level table (for a
 * logical table, the member it starts in) and its first row. Every upright
 * box records the order of its first fragment (UprightBlockPageIndex.boxOrders). */
function uprightFragmentOrder(request: BodyTableRequest): number {
  const cursor = request.cursor;
  const tableIndex = request.input.source.path[0]!
    + (cursor?.kind === 'adjacent-table-group' ? cursor.cursor.tableIndex : 0);
  const rowIndex = cursor === undefined ? 0
    : cursor.kind === 'table' ? cursor.cursor.rowIndex : cursor.cursor.sourceRowIndex;
  return tableIndex * 2 ** 24 + rowIndex;
}

/** The registration of a placed block of positive size; an empty one
 * registers nothing. */
function uprightBlockRegistrationOf(
  box: LayoutRect | null,
  boxOrder: number,
  slots: UprightBlockRegistration['slots'],
  exclusions: UprightBlockRegistration['exclusions'],
  tail?: NonNullable<UprightBlockRegistration['tail']>,
): UprightBlockRegistration | null {
  const sized = (rect: LayoutRect) => rect.widthPt > 0 && rect.heightPt > 0;
  const kept = exclusions.filter((exclusion) => sized(exclusion.bounds));
  const keptBox = box && sized(box) ? box : null;
  if (!keptBox && slots.length === 0 && kept.length === 0) return null;
  return Object.freeze({
    box: keptBox,
    boxOrder,
    slots: Object.freeze([...slots]),
    exclusions: Object.freeze(kept),
    ...(tail ? { tail: Object.freeze(tail) } : {}),
  });
}

/** Brings the page's index of wrapped anchored drawings up to date with its
 * wrap rectangles, which only grow while a page is laid out: each is looked
 * at once. A wrapNone drawing excludes no text and is not recorded
 * (paragraphWrapExclusions); neither is a positioned table, which a vertical
 * section lays out in its rotated frame, not upright. A drawing registered
 * when the page's layout started is recorded below the tables placed since
 * for tables before its paragraph, as a table placed later moves the
 * drawings it covers (moveUprightPageDrawingsBelowTableBox): every table is
 * placed against the drawings where they lie. */
function syncUprightDrawings(state: BodyAcquisitionState, story: SourceRef): UprightBlockPageIndex {
  const index = uprightBlockIndexOf(state);
  const floats = state.floats;
  if (index.drawingsSeen > floats.length) {
    // The rectangles were replaced, not extended: start over.
    index.drawingCells.clear();
    index.drawingCellRange = null;
    index.wideDrawings.length = 0;
    index.drawingsSeen = 0;
  }
  for (let at = index.drawingsSeen; at < floats.length; at += 1) {
    const float = floats[at]!;
    if (float.kind !== 'shape' || !float.anchorOccurrenceId || float.authoredWrap === undefined) continue;
    const anchorIndex = anchorOccurrenceBlockIndex(float.anchorOccurrenceId, story);
    if (anchorIndex === null) continue;
    const registeredFrame = Object.freeze({
      xPt: float.imageX, yPt: float.imageY, widthPt: float.imageW, heightPt: float.imageH,
    });
    addUprightPageDrawing(index, {
      anchorIndex,
      exclusion: Object.freeze({
        xPt: float.xLeft, yPt: float.yTop, widthPt: float.xRight - float.xLeft, heightPt: float.yBottom - float.yTop,
      }),
      frame: registeredFrame,
      registeredFrame,
    }, uprightDrawingLimits(state));
  }
  index.drawingsSeen = floats.length;
  return index;
}

/** The exclusions of the wrapped drawings anchored by a paragraph before the
 * table that lie in the way of a block spanning [`blockStartPt`,
 * `blockEndPt`] along the block axis and (`inlineFromPt`, `inlineToPt`)
 * along the inline axis: the page's drawings and `pending`, those of the
 * paragraph fragment the table is planned with, which are not registered
 * yet (BodyTableAcquisitionInput.precedingDrawings). An upright block table
 * meets them as it meets a table placed before it, without the table
 * clearance, and stacks below them
 * (`word-vertical-section-upright-block-table`). A drawing anchored after the
 * table is not met: it is the drawing that lies below the table
 * (uprightBlockCoversLaterDrawing). */
function uprightBlockObstaclesMet(
  index: UprightBlockPageIndex,
  tableIndex: number,
  pending: UprightBlockPageIndex | null,
  blockStartPt: number,
  blockEndPt: number,
  inlineFromPt: number,
  inlineToPt: number,
): LayoutRect[] {
  const met: LayoutRect[] = [];
  const inTheWay = (box: LayoutRect) => box.xPt < inlineToPt && box.xPt + box.widthPt > inlineFromPt
    && box.yPt < blockEndPt && blockStartPt < box.yPt + box.heightPt;
  for (const drawing of uprightPageDrawingsNear(index, blockStartPt, blockEndPt, inlineFromPt, inlineToPt)) {
    if (drawing.anchorIndex < tableIndex && inTheWay(drawing.exclusion)) met.push(drawing.exclusion);
  }
  if (pending) {
    for (const drawing of uprightPageDrawingsNear(pending, blockStartPt, blockEndPt, inlineFromPt, inlineToPt)) {
      if (inTheWay(drawing.exclusion)) met.push(drawing.exclusion);
    }
  }
  return met;
}

/** The drawings handed to a placement that are not registered yet
 * (BodyTableAcquisitionInput.precedingDrawings), in the two-axis cells the
 * page's drawings are kept in: built once for a list, so every window a
 * search asks for reads the drawings near it, not the whole list. */
const pendingDrawingIndexes = new WeakMap<readonly LayoutRect[], UprightBlockPageIndex>();
function pendingDrawingIndex(
  pending: readonly LayoutRect[] | undefined,
  limits: UprightDrawingLimits,
): UprightBlockPageIndex | null {
  if (!pending || pending.length === 0) return null;
  let index = pendingDrawingIndexes.get(pending);
  if (!index) {
    index = emptyUprightBlockPageIndex();
    for (let at = 0; at < pending.length; at += 1) {
      const rect = pending[at]!;
      addUprightPageDrawing(index, { anchorIndex: -1, exclusion: rect, frame: rect, registeredFrame: rect }, limits);
    }
    pendingDrawingIndexes.set(pending, index);
  }
  return index;
}

/**
 * `word-vertical-section-upright-block-table`: where a block with the given
 * box (`inlineExtentPt` along the column, `boxBlockExtentPt` between the
 * outer edges of its border rules) is placed when the flow reaches it at
 * `request.location`, against the boxes already on the page.
 *
 * The block starts at the current block position. Along the inline axis it
 * takes the first position from the column top at which its box clears the
 * boxes it meets (an earlier box is met up to
 * WORD_UPRIGHT_TABLE_TABLE_CLEARANCE_PT past its far edge), so overlapping
 * tables stack. Wrapped drawings anchored before the table are met where
 * they lie in the block's way (uprightBlockObstaclesMet).
 * A multi-column section never splits or moves a block: one that would pass
 * the page's bottom edge from the column top is first raised to end there,
 * no higher than the page top, and the search starts from that raised
 * position, so it is never raised into an earlier box.
 *
 * `first` (the first fragment of a table, with the tables that run on from
 * it): a block that does not fit below the boxes it meets clears the nearest
 * of them along the block axis instead. In a one-column section that holds
 * for a block that fits the band from the column top (a taller one
 * paginates); a multi-column section lets it run to the bottom page edge
 * below a table, and to the end of its column below a drawing, past which
 * it lies the text clearance from the drawing.
 * Later fragments and owner segments stay at the flow's block position
 * (library policy). `from` starts the search at a given position instead of
 * the flow's: rows that continue the table placed just before them.
 *
 * The result depends only on the request and the boxes placed before, and
 * every box it returns clears all of them: each pass over the met boxes moves
 * the block past every box it overlaps, and each move along the block axis
 * leaves at least the nearest met box behind for good, so the search ends
 * after at most as many moves as there are boxes on the page.
 */
function searchUprightBlock(
  index: UprightBlockPageIndex,
  physical: PhysicalAnchorFrame,
  request: BodyTableRequest,
  first: boolean,
  inlineExtentPt: number,
  boxBlockExtentPt: number,
  from?: Readonly<{ blockStartPt: number; inlineStartPt: number }>,
): Readonly<{
  blockStartPt: number;
  inlineStartPt: number;
  /** Inline extent from the block's inline start to the end of the page's body band. */
  inlineRoomPt: number;
  /** Inline extent of the page's body band from the column top. */
  inlineBandPt: number;
  /** The block was moved along the block axis past the block extent left. */
  leavesBlockExtent: boolean;
}> {
  const cursorPt = request.location.cursorPt;
  const multiColumn = request.location.section.columns.length > 1;
  const bandEndPt = physical.pageHeight - physical.marginBottom;
  const inlineOriginPt = from ? from.inlineStartPt : multiColumn
    ? Math.max(0, Math.min(cursorPt.xPt, physical.pageHeight - inlineExtentPt))
    : cursorPt.xPt;
  const flowBlockPt = from ? from.blockStartPt : cursorPt.yPt;
  const pendingDrawings = pendingDrawingIndex(request.precedingDrawings, {
    inlineEndPt: physical.pageHeight, blockEndPt: Number.POSITIVE_INFINITY,
    tableClearancePt: WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT,
  });
  let blockStartPt = flowBlockPt;
  let inlineStartPt = inlineOriginPt;
  for (;;) {
    const blockEndPt = blockStartPt + boxBlockExtentPt;
    const met = uprightBlockBoxesMet(
      index, blockStartPt, blockEndPt, WORD_UPRIGHT_TABLE_TABLE_CLEARANCE_PT, uprightFragmentOrder(request),
    );
    let nearestEndPt = Number.POSITIVE_INFINITY;
    for (const box of met) {
      nearestEndPt = Math.min(nearestEndPt, box.yPt + box.heightPt + WORD_UPRIGHT_TABLE_TABLE_CLEARANCE_PT);
    }
    // The first inline position from the origin at which the block clears
    // the met boxes and the drawings in its way. The met boxes are ordered
    // by inline start and passed once: a box passed without overlap ends at
    // or before the block's start, which only grows, and one starting past
    // the block's end is kept for the next round. The drawings are asked for
    // in the block's current inline window only; any position clear of one
    // in the window lies past its end, so the block moves there and the
    // window is asked again. Each round passes at least one box or drawing.
    inlineStartPt = inlineOriginPt;
    let next = 0;
    let metDrawing = false;
    for (;;) {
      while (next < met.length && met[next]!.xPt < inlineStartPt + inlineExtentPt) {
        const box = met[next]!;
        inlineStartPt = Math.max(inlineStartPt, box.xPt + box.widthPt);
        next += 1;
      }
      const obstacles = uprightBlockObstaclesMet(
        index, request.input.source.path[0]!, pendingDrawings,
        blockStartPt, blockEndPt, inlineStartPt, inlineStartPt + inlineExtentPt,
      );
      if (obstacles.length === 0) break;
      metDrawing = true;
      for (const box of obstacles) {
        // Past a drawing a table of a multi-column section lies the text
        // clearance from it.
        nearestEndPt = Math.min(
          nearestEndPt, box.yPt + box.heightPt + (multiColumn ? WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT : 0),
        );
        inlineStartPt = Math.max(inlineStartPt, box.xPt + box.widthPt);
      }
    }
    // A multi-column section lets a block below a table run to the bottom
    // page edge, but one below a drawing only to the end of its column.
    const fitsBelow = inlineStartPt + inlineExtentPt <= (!multiColumn ? bandEndPt
      : metDrawing ? cursorPt.xPt + request.availableInlineExtentPt : physical.pageHeight);
    if (!first || inlineStartPt === inlineOriginPt || fitsBelow
      || (!multiColumn && inlineOriginPt + inlineExtentPt > bandEndPt)) break;
    blockStartPt = nearestEndPt;
  }
  return Object.freeze({
    blockStartPt,
    inlineStartPt,
    inlineRoomPt: bandEndPt - inlineStartPt,
    inlineBandPt: bandEndPt - cursorPt.xPt,
    // Against the block extent the request gives the block, which can be
    // less than the location's (the footnotes a table refers to).
    leavesBlockExtent: blockStartPt !== flowBlockPt
      && blockStartPt + boxBlockExtentPt > Math.min(
        request.location.availableBounds.yPt + request.location.availableBounds.heightPt,
        request.location.cursorPt.yPt + (request.availableBlockExtentPt ?? Number.POSITIVE_INFINITY),
      ),
  });
}

/** One table of a run of upright block tables. */
type UprightRunMember = Readonly<{
  /** Body index of the table, or of the first table of a logical table. */
  sourceIndex: number;
  /** Body index of its last table. */
  lastSourceIndex: number;
  source: SourceRef;
  retained: RetainedTableAcquisition;
  gridExtentPt: number;
  edges: UprightBlockEdges;
  inlineExtentPt: number;
  boxBlockExtentPt: number;
}>;

function uprightRunMember(
  physical: PhysicalAnchorFrame,
  source: SourceRef,
  retained: RetainedTableAcquisition,
  lastSourceIndex = source.path[0]!,
): UprightRunMember {
  const gridExtentPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const edges = uprightBlockEdgeOverhangsPt(retained.layout, gridExtentPt, physical);
  return Object.freeze({
    sourceIndex: source.path[0]!,
    lastSourceIndex,
    source,
    retained,
    gridExtentPt,
    edges,
    inlineExtentPt: retained.layout.advancePt,
    boxBlockExtentPt: edges.nearPt + gridExtentPt + edges.farPt,
  });
}

/** Whether a table with this box starts a fresh region because its box (grid
 * and border rules) is wider than the block extent left while a fresh page
 * has more. That holds in the last column of a page, and so in a one-column
 * section; in an earlier column of a multi-column section the table stays
 * where the flow reached it, overflowing. */
function uprightBlockTooWide(request: BodyTableRequest, boxBlockExtentPt: number): boolean {
  const multiColumn = request.location.section.columns.length > 1;
  return (!multiColumn || request.lastColumnOfPage === true)
    && boxBlockExtentPt > request.availableBlockExtentPt
    && request.availableBlockExtentPt < request.freshPageBlockExtentPt;
}

/** The outcome of placing the first fragment of an upright block table the
 * flow reaches, with the tables that run on from it. */
/** The footnotes a table refers to, read from its acquisition once. */
const uprightTableFootnoteIdsByAcquisition = new WeakMap<RetainedTableAcquisition, readonly string[]>();
function uprightTableFootnoteIds(retained: RetainedTableAcquisition): readonly string[] {
  let ids = uprightTableFootnoteIdsByAcquisition.get(retained);
  if (!ids) uprightTableFootnoteIdsByAcquisition.set(retained, ids = footnoteIdsInRetainedSlice(retained.layout));
  return ids;
}

type UprightRunPlan =
  /** The table starts a fresh column or page instead. */
  | Readonly<{ kind: 'fresh-region'; freshPage: boolean }>
  /** A one-column section paginates the table by rows from this position. */
  | Readonly<{ kind: 'rows'; found: ReturnType<typeof searchUprightBlock> }>
  /** `slots[0]` is the table's, the others those of the tables after it
   * that run on from it on this page; `box` bounds the tables placed whole
   * (the last slot can be one that splits by rows and is not in it). */
  | Readonly<{
    kind: 'placed'; slots: readonly UprightBlockSlot[]; box: LayoutRect; lastSourceIndex: number;
    /** The footnotes the placed tables refer to. */
    footnoteReferenceIds: readonly string[];
  }>;

/** The synthetic inputs of the logical tables a run was taken with, by layout
 * session and first body index: one object per logical table, so its
 * combined acquisition is made once (combinedAdjacentGroupAcquisition). */
const uprightGroupInputsBySession = new WeakMap<object, Map<number, AdjacentBodyTableRequest['input']>>();

/** The logical tables acquired upright, by layout session and first body
 * index, with the extents they were acquired for and whether their members
 * are reusable across pages. */
const uprightLogicalTablesBySession = new WeakMap<object, Map<number, Readonly<{
  retained: RetainedTableAcquisition;
  lastSourceIndex: number;
  reusableAcrossPages: boolean;
  availableInlineExtentPt: number;
  freshPageBlockExtentPt: number;
}>>>();

/**
 * The upright table that starts at body index `index` in a vertical section:
 * a single table, or the tables the parser grouped into one logical table
 * (ECMA-376 §17.4.37), acquired as that one table. Null when no upright
 * table starts there (not a table, or positioned). `input` is the
 * paginator's input for it, when the flow has reached it.
 */
function uprightTableAcquisitionAt(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  index: number,
  input?: BodyTableRequest['input'],
): Readonly<{ source: SourceRef; retained: RetainedTableAcquisition; lastSourceIndex: number }> | null {
  const { dependencies, state, effectiveTablePositioning } = context;
  const body = dependencies.source.blocks.body;
  const element = body[index];
  if (element?.type !== 'table') return null;
  const memberState = bodyTableAcquisitionState(state, element, effectiveTablePositioning);
  if (memberState.uprightPhysicalTable !== true) return null;
  const source = input?.source ?? Object.freeze({ ...request.input.source, path: Object.freeze([index]) });
  const sequenceId = state.acquisitionInputs.tableFormatInput(element).logicalSequenceId ?? null;
  let lastIndex = index;
  if (input?.kind === 'adjacent-table-group') lastIndex = index + input.tables.length - 1;
  else if (!input && sequenceId !== null) {
    for (let next = index + 1; ; next += 1) {
      const following = body[next];
      if (following?.type !== 'table'
        || state.acquisitionInputs.tableFormatInput(following).logicalSequenceId !== sequenceId) break;
      lastIndex = next;
    }
  }
  if (lastIndex === index) {
    operations.computeTablePtLayout(memberState, element, request.availableInlineExtentPt, index);
    return { source, retained: retainedTableRecord(state, index).acquisition, lastSourceIndex: index };
  }
  // A logical table continued over pages is asked for on each of them. One
  // whose members hold nothing page-dependent is acquired once. One whose
  // members do is acquired again only where its first fragment is placed:
  // a continued fragment takes its rows from that acquisition, and the
  // page-dependent blocks of the rows it takes are acquired again on the
  // page that takes them (takeUprightTableRows), so each continuation page
  // reacquires the content of the members that land on it and nothing else.
  let logicalTables = uprightLogicalTablesBySession.get(context.sessionState);
  if (!logicalTables) uprightLogicalTablesBySession.set(context.sessionState, logicalTables = new Map());
  const reusable = logicalTables.get(index);
  if (reusable && reusable.lastSourceIndex === lastIndex
    && (reusable.reusableAcrossPages || request.cursor !== undefined)
    && reusable.availableInlineExtentPt === request.availableInlineExtentPt
    && reusable.freshPageBlockExtentPt === request.freshPageBlockExtentPt) {
    return { source, retained: reusable.retained, lastSourceIndex: lastIndex };
  }
  const members: Array<Readonly<{ table: typeof element; sourceIndex: number }>> = [];
  for (let at = index; at <= lastIndex; at += 1) {
    const table = body[at];
    if (table?.type !== 'table') throw new Error('Logical table member is not a table');
    members.push({ table, sourceIndex: at });
  }
  let groupInput = input?.kind === 'adjacent-table-group' ? input : undefined;
  if (!groupInput) {
    let inputs = uprightGroupInputsBySession.get(context.sessionState);
    if (!inputs) uprightGroupInputsBySession.set(context.sessionState, inputs = new Map());
    groupInput = inputs.get(index);
    if (!groupInput) {
      groupInput = Object.freeze({
        kind: 'adjacent-table-group' as const,
        logicalSequenceId: sequenceId!,
        source,
        tables: Object.freeze(members.map((member) => Object.freeze({
          kind: 'table' as const,
          source: Object.freeze({ ...source, path: Object.freeze([member.sourceIndex]) }),
          rowCount: member.table.rows.length,
        }))),
      });
      inputs.set(index, groupInput);
    }
  }
  const records = operations.computeAdjacentTablePtLayouts(memberState, members, request.availableInlineExtentPt);
  const retained = combinedAdjacentGroupAcquisition(
    context, { ...request, input: groupInput } as AdjacentBodyTableRequest, operations, records,
  );
  let reusableAcrossPages = true;
  for (const member of members) {
    if (!retainedTableRecord(state, member.sourceIndex).reusableAcrossPages) {
      reusableAcrossPages = false;
      break;
    }
  }
  logicalTables.set(index, {
    retained,
    lastSourceIndex: lastIndex,
    reusableAcrossPages,
    availableInlineExtentPt: request.availableInlineExtentPt,
    freshPageBlockExtentPt: request.freshPageBlockExtentPt,
  });
  return { source, retained, lastSourceIndex: lastIndex };
}

/**
 * The upright block table that starts at body index `index`, as one member
 * of a run (`word-vertical-section-upright-block-table`). Null when none
 * starts there: no upright table, one without rows, or one projected into
 * owner segments, which is placed segment by segment instead.
 */
function uprightUnitAt(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  physical: PhysicalAnchorFrame,
  index: number,
  input?: BodyTableRequest['input'],
): UprightRunMember | null {
  const acquired = uprightTableAcquisitionAt(context, request, operations, index, input);
  if (!acquired || acquired.retained.input.rows.length === 0
    || tableOwnerSegments(acquired.retained.input)) return null;
  return uprightRunMember(physical, acquired.source, acquired.retained, acquired.lastSourceIndex);
}

/** A run of upright block tables that directly follow one another in the
 * body, from the table the flow reached first: its tables in order, the
 * inline extent of each table and all after it, and the box extent of the
 * widest. Updated in place when a table measures differently on its page. */
type UprightRun = {
  readonly availableInlineExtentPt: number;
  readonly writingMode: PhysicalAnchorFrame['writingMode'];
  readonly members: UprightRunMember[];
  /** `suffixInlinePt[i]`: inline extent of members `i` and later. */
  readonly suffixInlinePt: Float64Array;
  boxBlockExtentPt: number;
};

/** By layout session: a later session can measure the same tables differently. */
const uprightRunsBySession = new WeakMap<object, Map<number, Readonly<{ run: UprightRun; at: number }>>>();

/** Whether a table now measures as the run recorded it. */
function uprightRunMemberHolds(recorded: UprightRunMember, current: UprightRunMember): boolean {
  return recorded.inlineExtentPt === current.inlineExtentPt
    && recorded.boxBlockExtentPt === current.boxBlockExtentPt;
}

/** Records how the run's table `at` measures now. The inline extents from
 * `fromAt` (a table at or before it, the first the current page still
 * reads) are brought up to date; nothing is measured, and tables placed on
 * earlier pages keep what was recorded for them. */
function refreshUprightRunMember(run: UprightRun, at: number, current: UprightRunMember, fromAt: number): void {
  run.members[at] = current;
  for (let index = at; index >= fromAt; index -= 1) {
    run.suffixInlinePt[index] = run.suffixInlinePt[index + 1]! + run.members[index]!.inlineExtentPt;
  }
  run.boxBlockExtentPt = Math.max(run.boxBlockExtentPt, current.boxBlockExtentPt);
}

/**
 * The run of directly following upright block tables that holds `leader`,
 * the table the flow reached, and its place in it. A run is laid out as one
 * table over however many pages it takes, so its tables are acquired once,
 * when the flow first reaches one of them, and every later page reads their
 * dimensions instead of walking the rest of the run again.
 *
 * A table whose cells hold page-dependent content (a PAGE field, say) can
 * measure differently on the page it is placed on than when the run was
 * taken. `leader` is how it measures now: where that differs, only its own
 * record is replaced. Nothing else is measured again; the tables of earlier
 * pages are placed and stay as they are, and each later table is checked on
 * the page that places it (planUprightRun).
 */
function uprightRunAt(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  physical: PhysicalAnchorFrame,
  leader: UprightRunMember,
): Readonly<{ run: UprightRun; at: number }> {
  let runs = uprightRunsBySession.get(context.sessionState);
  if (!runs) uprightRunsBySession.set(context.sessionState, runs = new Map());
  const cached = runs.get(leader.sourceIndex);
  if (cached && cached.run.availableInlineExtentPt === request.availableInlineExtentPt
    && cached.run.writingMode === physical.writingMode) {
    if (!uprightRunMemberHolds(cached.run.members[cached.at]!, leader)) {
      refreshUprightRunMember(cached.run, cached.at, leader, cached.at);
    }
    return cached;
  }
  const members = [leader];
  for (let index = leader.lastSourceIndex + 1; ;) {
    const member = uprightUnitAt(context, request, operations, physical, index);
    if (!member) break;
    members.push(member);
    index = member.lastSourceIndex + 1;
  }
  const suffixInlinePt = new Float64Array(members.length + 1);
  let boxBlockExtentPt = 0;
  for (let index = members.length - 1; index >= 0; index -= 1) {
    suffixInlinePt[index] = suffixInlinePt[index + 1]! + members[index]!.inlineExtentPt;
    boxBlockExtentPt = Math.max(boxBlockExtentPt, members[index]!.boxBlockExtentPt);
  }
  const run: UprightRun = {
    availableInlineExtentPt: request.availableInlineExtentPt,
    writingMode: physical.writingMode,
    members,
    suffixInlinePt,
    boxBlockExtentPt,
  };
  members.forEach((member, at) => runs.set(member.sourceIndex, Object.freeze({ run, at })));
  return runs.get(leader.sourceIndex)!;
}

/** Block start of a table's box inside its run's box, which starts at
 * `runBlockStartPt`: the tables of a run share the physical left edge, the
 * far edge of the box in the clockwise frame and the near edge in
 * vertical-lr. */
function uprightRunMemberBlockStartPt(
  physical: PhysicalAnchorFrame,
  run: UprightRun,
  runBlockStartPt: number,
  boxBlockExtentPt: number,
): number {
  return physical.writingMode === 'vertical-lr'
    ? runBlockStartPt : runBlockStartPt + run.boxBlockExtentPt - boxBlockExtentPt;
}

/**
 * `word-vertical-section-upright-block-table`: places the upright block table
 * `leader` the flow reaches at `request.location`, together with the upright
 * block tables that directly follow it in the body with nothing between them.
 *
 * Under that rule such tables are laid out as one table: the rows of each
 * start where the rows of the one before end, their left edges aligned, and the whole takes
 * the positions, the pagination and the exclusion of a single table as wide
 * as its widest member. So the run is placed by its total inline extent and
 * its widest box, and its narrower tables lie at the physical left of that
 * box (the far edge in the clockwise frame, the near edge in vertical-lr).
 *
 * A multi-column section never splits or moves a table: the whole run is one
 * block, raised together from the bottom page edge. In a one-column section
 * a run taller than the inline extent below its start starts a fresh page
 * when content precedes it; on a fresh page its tables fill the band whole
 * until one does not fit, which is cut between rows there and continues on
 * the next page, where the tables after it follow its last rows
 * (`continuing`).
 *
 * Library policy: a run whose tables differ in their border rules is aligned
 * by the tables' boxes.
 */
function planUprightRun(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  physical: PhysicalAnchorFrame,
  leader: UprightRunMember,
  reserving: boolean,
  /** Body indices of the run's tables already measured on this page for
   * this plan (replanIfStale). */
  checked: Set<number> = new Set(),
): UprightRunPlan {
  const { state } = context;
  const index = syncUprightDrawings(state, request.input.source);
  const flowDomainId = request.location.flowDomainId;
  const multiColumn = request.location.section.columns.length > 1;
  const pageHasContent = request.pageHasContent === true;
  const { run, at } = uprightRunAt(context, request, operations, physical, leader);
  const { members } = run;
  const boxBlockExtentPt = run.boxBlockExtentPt;
  const totalInlinePt = run.suffixInlinePt[at]!;
  // The tables a plan places on this page have to measure here as the run
  // recorded them. Where one differs its record is replaced and the plan
  // made anew, which can admit tables the first plan did not hold; those are
  // checked in turn. Each table is measured once on its page (`checked`), so
  // the checks grow with the tables, not their square, and the plan is made
  // at most once more per table it places.
  const replanIfStale = (plan: UprightRunPlan): UprightRunPlan => {
    if (plan.kind !== 'placed') return plan;
    let changed = false;
    for (let offset = 1; offset < plan.slots.length; offset += 1) {
      const recorded = members[at + offset]!;
      if (checked.has(recorded.sourceIndex)) continue;
      checked.add(recorded.sourceIndex);
      const current = uprightUnitAt(context, request, operations, physical, recorded.sourceIndex);
      if (!current) throw new Error('Upright block table run member is no longer an upright table');
      if (uprightRunMemberHolds(recorded, current)) continue;
      refreshUprightRunMember(run, at + offset, current, at);
      changed = true;
    }
    return changed ? planUprightRun(context, request, operations, physical, leader, reserving, checked) : plan;
  };
  // `wholeCount` tables from this one placed whole from `inlineStartPt`,
  // then at most one cut by rows.
  const placedOf = (blockStartPt: number, inlineStartPt: number, wholeCount: number): UprightRunPlan => {
    const slots: UprightBlockSlot[] = [];
    let inlineEndPt = inlineStartPt;
    const slotOf = (member: UprightRunMember, splitRows: boolean): UprightBlockSlot => Object.freeze({
      sourceIndex: member.sourceIndex,
      lastSourceIndex: member.lastSourceIndex,
      first: member === members[at],
      flowDomainId,
      bounds: Object.freeze({
        xPt: inlineEndPt,
        yPt: uprightRunMemberBlockStartPt(physical, run, blockStartPt, member.boxBlockExtentPt),
        widthPt: member.inlineExtentPt,
        heightPt: member.boxBlockExtentPt,
      }),
      reserved: reserving,
      anyColumn: false,
      pinned: false,
      splitRows,
    });
    for (let offset = 0; offset < wholeCount; offset += 1) {
      const member = members[at + offset]!;
      slots.push(slotOf(member, false));
      inlineEndPt += member.inlineExtentPt;
    }
    const box: LayoutRect = Object.freeze({
      xPt: inlineStartPt, yPt: blockStartPt, widthPt: inlineEndPt - inlineStartPt, heightPt: boxBlockExtentPt,
    });
    if (at + wholeCount < members.length) slots.push(slotOf(members[at + wholeCount]!, true));
    // The notes of the tables placed whole: a table cut by rows here keeps
    // only the rows the band holds, which its own row pagination admits the
    // notes of.
    const footnoteReferenceIds: string[] = [];
    for (let offset = 0; offset < wholeCount; offset += 1) {
      const ids = uprightTableFootnoteIds(members[at + offset]!.retained);
      for (let id = 0; id < ids.length; id += 1) footnoteReferenceIds.push(ids[id]!);
    }
    return Object.freeze({
      kind: 'placed' as const, slots, box, lastSourceIndex: members[at + wholeCount - 1]!.lastSourceIndex,
      footnoteReferenceIds: Object.freeze(footnoteReferenceIds),
    });
  };
  // Rows that continue the upright table placed just before this one in the
  // same column start where it ends and are cut where the band ends.
  const tail = index.tail;
  const continuing = !multiColumn && tail !== null && tail.sourceIndex === leader.sourceIndex - 1
    && tail.flowDomainId === flowDomainId;
  if (uprightBlockTooWide(request, boxBlockExtentPt)) return Object.freeze({ kind: 'fresh-region', freshPage: false });
  const found = searchUprightBlock(
    index, physical, request, !continuing, totalInlinePt, boxBlockExtentPt,
    continuing ? { blockStartPt: tail.bounds.yPt, inlineStartPt: tail.bounds.xPt + tail.bounds.widthPt } : undefined,
  );
  if (found.leavesBlockExtent && pageHasContent) return Object.freeze({ kind: 'fresh-region', freshPage: true });
  if (multiColumn || totalInlinePt <= found.inlineRoomPt) {
    return replanIfStale(placedOf(found.blockStartPt, found.inlineStartPt, members.length - at));
  }
  // Too tall for the inline extent below its start: a fresh page when
  // content precedes it, unless it continues the table before it.
  if (pageHasContent && !continuing) return Object.freeze({ kind: 'fresh-region', freshPage: true });
  // The tables that fit whole, summed in body order. That sum and the run's
  // recorded total are taken in different orders and can differ by rounding,
  // so a run the total calls too tall can still fit to its last table here.
  let wholeCount = 0;
  let filledPt = 0;
  while (at + wholeCount < members.length
    && filledPt + members[at + wholeCount]!.inlineExtentPt <= found.inlineRoomPt) {
    filledPt += members[at + wholeCount]!.inlineExtentPt;
    wholeCount += 1;
  }
  // A plan for the paragraph before the tables holds whole tables only.
  if (wholeCount === 0 || reserving) return reserving
    ? Object.freeze({ kind: 'fresh-region', freshPage: true })
    : Object.freeze({ kind: 'rows', found });
  return replanIfStale(placedOf(found.blockStartPt, found.inlineStartPt, wholeCount));
}

/** `word-vertical-section-upright-block-table`: the top-level upright block
 * table of `request` (a table, or the tables of one logical table), or null
 * when it is none or is projected into owner segments. */
function uprightBlockLeader(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
): Readonly<{ leader: UprightRunMember; physical: PhysicalAnchorFrame }> | null {
  const physical = context.state.verticalPhys;
  if (!physical || !uprightBlockPaginates(physical)) return null;
  const leader = uprightUnitAt(
    context, request, operations, physical, request.input.source.path[0]!, request.input,
  );
  return leader ? { leader, physical } : null;
}

/** The exclusion of the tables placed whole by a plan: one rectangle over
 * their common box, as of one table. */
function uprightRunExclusion(
  request: BodyTableRequest,
  box: LayoutRect,
): UprightBlockRegistration['exclusions'] {
  return [uprightBlockExclusion(request.input.source, tableFragmentStartKey(undefined), box)];
}

/** Whether a block placed at `box` from `request.location` lies over the
 * inline band of a column laid out before its own: a multi-column section
 * raised it above its column's top, and its column is not the first. Text
 * laid out there earlier has to flow around it too. */
function uprightBlockReachesEarlierColumn(request: BodyTableRequest, box: LayoutRect): boolean {
  const columns = request.location.section.columns;
  if (columns.length <= 1 || !(box.xPt < request.location.cursorPt.xPt)) return false;
  let firstColumnStartPt = Number.POSITIVE_INFINITY;
  for (const column of columns) firstColumnStartPt = Math.min(firstColumnStartPt, column.xPt);
  return request.location.cursorPt.xPt > firstColumnStartPt;
}

/** Whether a block placed at `box` covers a wrapped drawing of the page that
 * a paragraph after the table anchors (registered when the page's layout
 * started, before the table was placed). Such a drawing lies below the table
 * (`word-vertical-section-upright-block-table`), which its placement can only
 * honour once the table's box is known when the page's layout starts. */
function uprightBlockCoversLaterDrawing(state: BodyAcquisitionState, table: SourceRef, box: LayoutRect): boolean {
  const tableIndex = table.path[0]!;
  const index = syncUprightDrawings(state, table);
  // The drawing as the page's wrap rectangles hold it: once the index has
  // moved it below this table, the text laid out so far still wraps around
  // it where it was registered.
  for (const { anchorIndex, registeredFrame: frame } of uprightPageDrawingsNear(
    index, box.yPt, box.yPt + box.heightPt, box.xPt, box.xPt + box.widthPt,
  )) {
    if (anchorIndex > tableIndex
      && frame.xPt < box.xPt + box.widthPt && box.xPt < frame.xPt + frame.widthPt
      && frame.yPt < box.yPt + box.heightPt && box.yPt < frame.yPt + frame.heightPt) return true;
  }
  return false;
}

/** A plan's blocks as a page pin: every table in its slot for good, with the
 * exclusion, to be registered when the page's layout starts
 * (BodyLayoutSession.commitUprightBlocks). */
function uprightPagePin(
  request: BodyTableRequest,
  box: LayoutRect,
  slots: readonly UprightBlockSlot[],
  lastSourceIndex: number,
): UprightBlockRegistration {
  return Object.freeze({
    box,
    boxOrder: uprightFragmentOrder(request),
    slots: Object.freeze(slots.map((slot) =>
      Object.freeze({ ...slot, reserved: true, anyColumn: true, pinned: true }))),
    exclusions: Object.freeze(uprightRunExclusion(request, box)),
    tail: Object.freeze({ sourceIndex: lastSourceIndex, flowDomainId: request.location.flowDomainId, bounds: box }),
  });
}

/**
 * Keep-with-next look-ahead for an upright block table
 * (`word-vertical-section-upright-block-table`): whether the table of
 * `request` starts on this page if the flow reaches it at the request's
 * location, decided by the placement itself (planUprightRun) against the
 * blocks already on the page, so a table that placement moves to a fresh
 * page to clear an earlier one is not promised to the paragraph kept with
 * it; with the table's retained layout. Null when the table is not such a
 * block. Changes nothing.
 */
export function bodyUprightBlockLookAhead(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
): Readonly<{
  startsHere: boolean;
  layout: TableLayout;
  /** Where the box of the tables placed with it ends along the block axis,
   * when it is placed whole here. */
  blockEndPt?: number;
  /** The footnotes those tables refer to. */
  footnoteReferenceIds: readonly string[];
}> | null {
  const { services, state, sessionState } = context;
  const restore = sessionState.location;
  operations.setBodyAcquisitionLocation(services, state, sessionState, request.location);
  try {
    const found = uprightBlockLeader(context, request, operations);
    if (!found) return null;
    const layout = found.leader.retained.layout;
    const own = uprightTableFootnoteIds(found.leader.retained);
    const slot = uprightBlockIndexOf(state).slots.get(found.leader.sourceIndex);
    if (slot && (slot.anyColumn || slot.flowDomainId === request.location.flowDomainId)) {
      return {
        startsHere: true, layout, blockEndPt: slot.bounds.yPt + slot.bounds.heightPt, footnoteReferenceIds: own,
      };
    }
    const plan = planUprightRun(context, request, operations, found.physical, found.leader, false);
    return {
      startsHere: plan.kind !== 'fresh-region',
      layout,
      ...(plan.kind === 'placed' ? { blockEndPt: plan.box.yPt + plan.box.heightPt } : {}),
      footnoteReferenceIds: plan.kind === 'placed' ? plan.footnoteReferenceIds : own,
    };
  } finally {
    operations.setBodyAcquisitionLocation(services, state, sessionState, restore);
  }
}

/**
 * BodyLayoutSession.planUprightBlocks
 * (`word-vertical-section-upright-table-rewrapped-strip`): the upright block
 * tables placed whole on this page when the flow reaches the table of
 * `request` at its location, as slots the tables take when the flow does
 * reach them, with their text exclusion. The paginator asks this for the
 * paragraph fragment that ends just before the table, laid out without the
 * table, and lays that fragment out again around the result. Null when the
 * table is not placed whole there or already has a slot.
 */
export function planBodyUprightBlocks(
  context: BodyTableMeasurementContext,
  request: Parameters<NonNullable<BodyLayoutSession['planUprightBlocks']>>[0],
  operations: BodyTableMeasurementOperations,
): UprightBlockRegistration | null {
  const { services, state, sessionState } = context;
  const restore = sessionState.location;
  operations.setBodyAcquisitionLocation(services, state, sessionState, request.location);
  try {
    const found = uprightBlockLeader(context, request, operations);
    if (!found || request.cursor) return null;
    const { leader, physical } = found;
    if (uprightBlockIndexOf(state).slots.has(leader.sourceIndex)) return null;
    const plan = planUprightRun(context, request, operations, physical, leader, true);
    if (plan.kind !== 'placed') return null;
    return Object.freeze({
      box: plan.box,
      boxOrder: uprightFragmentOrder(request),
      slots: Object.freeze([...plan.slots]),
      exclusions: Object.freeze(uprightRunExclusion(request, plan.box)),
      tail: Object.freeze({ sourceIndex: plan.lastSourceIndex, flowDomainId: request.location.flowDomainId, bounds: plan.box }),
      ...(plan.footnoteReferenceIds.length > 0 ? { footnoteReferenceIds: plan.footnoteReferenceIds } : {}),
    });
  } finally {
    operations.setBodyAcquisitionLocation(services, state, sessionState, restore);
  }
}

/**
 * `word-vertical-section-upright-block-table`: one fragment of a top-level
 * upright block table of a vertical-rl or vertical-lr section frame.
 *
 * The block charges no flow advance: the following paragraph takes the next
 * line strip as if the block were absent. Its section-logical box, extended
 * by WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT past its far edge, is a text
 * exclusion met from its block start inclusive: every line strip meeting it
 * flows around it, which for a block at the column top means starting at its
 * inline end.
 *
 * Placement is decided once and never revised: a table with a slot (placed
 * with the paragraph before it, as one of the tables running on from an
 * earlier one, or pinned when the page's layout started) takes it; any other
 * first fragment is placed here with the tables that run on from it
 * (planUprightRun). A table placed ahead of the flow is where that plan put
 * it, whatever block position the flow has reached since
 * (`word-vertical-section-upright-table-rewrapped-strip`).
 *
 * `uprightPagePin`: the block lies over a column laid out before its own,
 * whose text has to flow around it as well, or over a wrapped drawing that a
 * later paragraph anchors, which has to lie below it. The paginator lays the
 * page out again with the pin registered from its start (body-paginator.ts).
 */
function measureUprightBlockTable(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  leader: UprightRunMember,
  cursor: ReturnType<typeof startTableFragmentCursor>,
  acquisitionState: BodyAcquisitionState,
  physical: PhysicalAnchorFrame,
): TableMeasureResult {
  const index = uprightBlockIndexOf(context.state);
  const flowDomainId = request.location.flowDomainId;
  const { retained, sourceIndex } = leader;
  const freshRegion = (freshPage: boolean) => Object.freeze({
    layout: retained.layout,
    blockExtentPt: 0,
    nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
    requiresFreshFlowRegion: true,
    ...(freshPage ? { requiresFreshPage: true } : {}),
  });
  const placed = (
    layout: TableLayout,
    bounds: LayoutRect,
    nextCursor: TableMeasureResult['nextCursor'],
    registration: UprightBlockRegistration | null,
    pin: UprightBlockRegistration | null = null,
  ): TableMeasureResult => {
    const origin = uprightBlockGridOrigin(physical, bounds, leader.gridExtentPt, leader.edges);
    return Object.freeze({
      layout,
      blockExtentPt: 0,
      nextCursor,
      ...(registration ? { uprightBlocks: registration } : {}),
      ...(pin ? { uprightPagePin: pin } : {}),
      uprightBlockEndPt: Math.max(
        bounds.yPt + bounds.heightPt, registration?.box ? registration.box.yPt + registration.box.heightPt : bounds.yPt,
      ),
      placement: Object.freeze({
        coordinateSpace: 'upright-physical' as const,
        xPt: origin.xPt,
        yPt: origin.yPt + layout.flowBounds.yPt,
        sectionFlowOwnership: 'host-flow' as const,
      }),
    });
  };
  const wholeAt = (bounds: LayoutRect) => takeUprightTableFragment(
    context, request, operations, retained,
    uprightBlockTranslationPt(
      uprightBlockGridOrigin(physical, bounds, leader.gridExtentPt, leader.edges), retained.layout,
    ),
    acquisitionState,
  ) as TableLayout;
  const { run } = uprightRunAt(context, request, operations, physical, leader);
  /** Whole rows from `cursor` that fit the band below `at`, the position of
   * the run's box; a row taller than the band takes a page of its own. Every
   * fragment keeps the box edges of the whole table (rows continued on a
   * later page stay offset by a rule only earlier rows carry) and its place
   * in the box of its run, whose extent it registers. */
  const rowsAt = (found: Readonly<{
    blockStartPt: number; inlineStartPt: number; inlineRoomPt: number; inlineBandPt: number;
  }>): TableMeasureResult => {
    // Rows pinned when the page's layout started (below) take the pinned
    // place, which is registered already.
    const segmentKey = uprightSegmentKey(sourceIndex, tableFragmentStartKey(request.cursor));
    const pinned = index.segmentPins.get(segmentKey);
    const at = pinned ? {
      ...found,
      blockStartPt: pinned.yPt,
      inlineStartPt: pinned.xPt,
      inlineRoomPt: physical.pageHeight - physical.marginBottom - pinned.xPt,
    } : found;
    const searched: LayoutRect = {
      xPt: at.inlineStartPt,
      yPt: uprightRunMemberBlockStartPt(physical, run, at.blockStartPt, leader.boxBlockExtentPt),
      widthPt: leader.inlineExtentPt,
      heightPt: leader.boxBlockExtentPt,
    };
    const taken = takeUprightTableRows(
      context, request, operations, wholeRowAcquisition(retained),
      uprightBlockTranslationPt(
        uprightBlockGridOrigin(physical, searched, leader.gridExtentPt, leader.edges), retained.layout,
      ),
      acquisitionState, cursor, at.inlineRoomPt, at.inlineBandPt,
    );
    if (!taken.fragment) {
      if (request.pageHasContent === true) return freshRegion(true);
      throw new Error('An upright table made no progress on a fresh page');
    }
    // The fragment lies inside the position found for the whole table.
    const bounds: LayoutRect = Object.freeze({
      ...searched,
      widthPt: Math.min(taken.fragment.advancePt, Math.max(0, at.inlineRoomPt)),
    });
    const box: LayoutRect = Object.freeze({
      xPt: bounds.xPt, yPt: at.blockStartPt, widthPt: bounds.widthPt, heightPt: run.boxBlockExtentPt,
    });
    const registration = pinned ? null : uprightBlockRegistrationOf(
      box, uprightFragmentOrder(request), [], [uprightBlockExclusion(request.input.source, tableFragmentStartKey(request.cursor), box)],
      { sourceIndex: leader.lastSourceIndex, flowDomainId, bounds: box },
    );
    // Rows that cover a wrapped drawing registered when the page's layout
    // started, which a paragraph after the table anchors, are pinned like
    // a whole table: the drawing and its exclusion lie below them in the
    // page's next layout, where they take this place again.
    return placed(
      taken.fragment as TableLayout,
      bounds,
      taken.nextCursor ? Object.freeze({ kind: 'table' as const, cursor: taken.nextCursor }) : null,
      registration,
      registration && uprightBlockCoversLaterDrawing(context.state, request.input.source, box)
        ? Object.freeze({ ...registration, segmentPin: Object.freeze({ key: segmentKey, bounds: box }) })
        : null,
    );
  };

  if (!request.cursor) {
    const slot = index.slots.get(sourceIndex);
    if (slot && (slot.anyColumn || slot.flowDomainId === flowDomainId)) {
      // A table that runs on from an earlier one still starts a fresh region
      // when the block extent left has shrunk below its box since.
      if (!slot.reserved && uprightBlockTooWide(request, leader.boxBlockExtentPt)) return freshRegion(false);
      if (slot.splitRows) {
        const bandEndPt = physical.pageHeight - physical.marginBottom;
        return rowsAt({
          blockStartPt: physical.writingMode === 'vertical-lr'
            ? slot.bounds.yPt : slot.bounds.yPt + slot.bounds.heightPt - run.boxBlockExtentPt,
          inlineStartPt: slot.bounds.xPt,
          inlineRoomPt: bandEndPt - slot.bounds.xPt,
          inlineBandPt: bandEndPt - request.location.cursorPt.xPt,
        });
      }
      // Placed with the paragraph before it and raised over an earlier
      // column: the tables of its plan become a page pin.
      let pin: UprightBlockRegistration | null = null;
      if (slot.first && slot.reserved && !slot.pinned) {
        // The tables placed with this one, and the box of them all.
        const slots = [slot];
        let box = slot.bounds;
        for (let next = slot.lastSourceIndex + 1; ;) {
          const follower = index.slots.get(next);
          if (!follower || follower.first || !follower.reserved
            || follower.flowDomainId !== slot.flowDomainId) break;
          slots.push(follower);
          const startPt = Math.min(box.yPt, follower.bounds.yPt);
          const endPt = Math.max(box.yPt + box.heightPt, follower.bounds.yPt + follower.bounds.heightPt);
          box = {
            xPt: box.xPt,
            yPt: startPt,
            widthPt: follower.bounds.xPt + follower.bounds.widthPt - box.xPt,
            heightPt: endPt - startPt,
          };
          next = follower.lastSourceIndex + 1;
        }
        if (uprightBlockReachesEarlierColumn(request, box)
          || uprightBlockCoversLaterDrawing(context.state, request.input.source, box)) {
          pin = uprightPagePin(request, Object.freeze(box), slots, slots.at(-1)!.lastSourceIndex);
        }
      }
      return placed(wholeAt(slot.bounds), slot.bounds, null, null, pin);
    }
    const plan = planUprightRun(context, request, operations, physical, leader, false);
    if (plan.kind === 'fresh-region') return freshRegion(plan.freshPage);
    if (plan.kind === 'placed') {
      const bounds = plan.slots[0]!.bounds;
      return placed(
        wholeAt(bounds), bounds, null,
        Object.freeze({
          box: plan.box,
          boxOrder: uprightFragmentOrder(request),
          slots: Object.freeze(plan.slots.slice(1)),
          exclusions: Object.freeze(uprightRunExclusion(request, plan.box)),
          tail: Object.freeze({ sourceIndex: plan.lastSourceIndex, flowDomainId, bounds: plan.box }),
          ...(plan.footnoteReferenceIds.length > 0 ? { footnoteReferenceIds: plan.footnoteReferenceIds } : {}),
        }),
        uprightBlockReachesEarlierColumn(request, plan.box)
          || uprightBlockCoversLaterDrawing(context.state, request.input.source, plan.box)
          ? uprightPagePin(request, plan.box, plan.slots, plan.lastSourceIndex) : null,
      );
    }
    return rowsAt(plan.found);
  }
  if (request.location.section.columns.length > 1) {
    throw new Error('An upright block table of a multi-column section must remain whole');
  }
  // Rows continued on a later page stay at the flow's block position and
  // clear the blocks they meet along the inline axis (library policy).
  return rowsAt(searchUprightBlock(
    syncUprightDrawings(context.state, request.input.source), physical, request, false,
    leader.inlineExtentPt, run.boxBlockExtentPt,
  ));
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
  const physical = context.state.verticalPhys;
  if (!physical) throw new Error('An upright table requires a vertical section');
  const physicalBandHeightPt = Math.max(
    retained.layout.advancePt,
    physical.pageHeight - physical.marginTop - physical.marginBottom,
  );
  const upright = takeUprightTableRows(
    context, request, operations, retained, translationPt, owner,
    startTableFragmentCursor(), physicalBandHeightPt, physicalBandHeightPt,
  );
  if (!upright.fragment || upright.nextCursor || upright.requiresFreshPage) {
    throw new Error('Upright table final-frame layout must remain atomic');
  }
  return upright.fragment;
}

/** The rows of a §17.6.20 upright physical table from `cursor` that fit a
 * physical band `availableHeightPt` tall (a fresh band: `freshHeightPt`),
 * in the physical page frame translated to `translationPt`. */
function takeUprightTableRows(
  context: BodyTableMeasurementContext,
  request: BodyTableRequest,
  operations: BodyTableMeasurementOperations,
  retained: RetainedTableAcquisition,
  translationPt: Readonly<{ xPt: number; yPt: number }>,
  owner: BodyAcquisitionState,
  cursor: ReturnType<typeof startTableFragmentCursor>,
  physicalBandHeightPt: number,
  freshHeightPt: number,
) {
  const { dependencies, services, state } = context;
  const physical = state.verticalPhys;
  if (!physical) throw new Error('An upright table requires a vertical section');
  const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const flowDomainId = `upright-physical-page:${request.location.pageIndex}`;
  const physicalMargin = {
    xPt: physical.marginLeft,
    yPt: physical.marginTop,
    widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
    heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
  };
  const physicalPage = { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight };
  return takeTableFragment(retained, cursor, {
    availableHeightPt: physicalBandHeightPt,
    freshPageHeightPt: freshHeightPt,
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
      occurrenceId: `${retained.input.id}:upright-page:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
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
  const { dependencies, state } = context;

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
  shareRotatedCellReacquisition(records.map((record) => record.input), combinedInput);
  const combinedLayout = layoutRetainedTableInput(
    combinedInput, adjacentGroupPlacement(request), services,
  ).layout;
  // ast-grep-ignore: no-plain-object-string-map -- keys are library-generated table flow ids.
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
  const { state, sessionState } = context;
  const { table, retained, cursor, pageHeightPt, authoredPositioning } = frame;

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
  const absoluteAnchorMustSplit =
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    retained.layout.advancePt > request.freshPageBlockExtentPt;
  // A page- or margin-anchored table is admitted against the page, as it is
  // clamped (resolveFloatingTableBoxPt); leaving the text area is decided
  // below, after placement (`word-page-anchored-table-text-area-move`).
  const admissionBlockEndPt = absoluteAnchorMustSplit
    ? request.location.availableBounds.yPt + request.location.availableBounds.heightPt
    : positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin'
      ? frames.page.yPt + frames.page.heightPt
      : request.location.availableBounds.yPt + request.location.availableBounds.heightPt;
  const freshAdmissionHeightPt = absoluteAnchorMustSplit
    ? request.freshPageBlockExtentPt
    : positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin'
      ? frames.page.heightPt
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
  // `word-page-anchored-table-text-area-move`: a page- or margin-anchored
  // table whose box ends below the bottom margin, or that never-displacement
  // moved above the top margin, moves to the next page together with the
  // flow after it to the next page (not the next column), unless nothing
  // precedes it on this page (moving would gain nothing there). Its
  // authored position applies on the next page.
  const box = resolved.placement.bounds;
  const textArea = frames.margin;
  if (
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    !isFloatingContinuation &&
    !absoluteAnchorMustSplit &&
    request.pageHasContent === true &&
    (box.yPt + box.heightPt > textArea.yPt + textArea.heightPt + FLOAT_OVERLAP_EPS ||
      (box.yPt < raw.y - FLOAT_OVERLAP_EPS && box.yPt < textArea.yPt - FLOAT_OVERLAP_EPS))
  ) {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'authored' as const,
      }),
      requiresFreshFlowRegion: true,
      requiresFreshPage: true,
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
  shareRotatedCellReacquisition(retained.input, input);
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
 * - an ordinary segment is an upright block of its own: it adds no flow
 *   advance and excludes text as an unsegmented upright table does, stays at
 *   the flow's block position and clears the blocks it meets along the
 *   inline axis (searchUprightBlock);
 * - a host resolves its carrier in the upright physical page frame — the
 *   frame this path already gives nested floating tables: physical page and
 *   margin, a text anchor's horizontal band is the physical margin box and its
 *   vertical band starts at the physical top of the current column
 *   (physicalInlineBandTopPt) — and advances the flow by 0. Its exclusion
 *   enters the logical body registry through the inverse of the section's
 *   upright projection (uprightPhysicalRectToLogical: clockwise, logical
 *   inline = physical y and logical block start = physical page width −
 *   physical right edge; tbLrV, logical block start = physical left edge; a
 *   section flow, its own counter-clockwise inverse), so
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
    // An ordinary segment of a vertical-rl or vertical-lr frame is fitted as
    // an upright block table is: by its box, grid and border rules, from the
    // last column of its page. A counter-clockwise flow fits its grid.
    const tooWide = uprightBlockPaginates(physical)
      ? uprightBlockTooWide(request, (() => {
          const edges = uprightBlockEdgeOverhangsPt(local.layout, gridWidthPt, physical);
          return edges.nearPt + gridWidthPt + edges.farPt;
        })())
      : gridWidthPt > request.availableBlockExtentPt
        && request.availableBlockExtentPt < request.freshPageBlockExtentPt;
    if (tooWide) {
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
    if (!uprightBlockPaginates(physical)) {
      const origin = uprightBlockPhysicalOrigin(
        physical, request.location.cursorPt, gridWidthPt, local.layout.advancePt,
      );
      const layout = takeUprightTableFragment(context, request, operations, local, origin, acquisitionOwner);
      return Object.freeze({
        layout,
        blockExtentPt: gridWidthPt,
        nextCursor,
        placement: Object.freeze({
          coordinateSpace: 'upright-physical' as const,
          xPt: origin.xPt + layout.flowBounds.xPt,
          yPt: origin.yPt + layout.flowBounds.yPt,
          sectionFlowOwnership: 'host-flow' as const,
        }),
      });
    }
    // An ordinary segment is an upright block of its own (library policy): it
    // stays at the flow's block position and clears the blocks it meets
    // along the inline axis, unless it was pinned when the page's layout
    // started, in which case it is where the pin has it.
    const edges = uprightBlockEdgeOverhangsPt(local.layout, gridWidthPt, physical);
    const boxExtentPt = edges.nearPt + gridWidthPt + edges.farPt;
    const segmentIndex = syncUprightDrawings(context.state, request.input.source);
    const fragmentStartKey = tableFragmentStartKey(request.cursor);
    const segmentKey = uprightSegmentKey(request.input.source.path[0]!, fragmentStartKey);
    const pinned = segmentIndex.segmentPins.get(segmentKey);
    const bounds: LayoutRect = pinned ?? (() => {
      const found = searchUprightBlock(
        segmentIndex, physical, request, false, local.layout.advancePt, boxExtentPt,
      );
      return Object.freeze({
        xPt: found.inlineStartPt, yPt: found.blockStartPt, widthPt: local.layout.advancePt, heightPt: boxExtentPt,
      });
    })();
    const origin = uprightBlockGridOrigin(physical, bounds, gridWidthPt, edges);
    const layout = takeUprightTableFragment(
      context, request, operations, local, uprightBlockTranslationPt(origin, local.layout), acquisitionOwner,
    );
    // A wrapped drawing anchored after the table lies below the segment
    // too, so its box enters the history of placed tables. Where such a
    // drawing was registered when the page's layout started, the segment is
    // pinned: its box, exclusion and place are registered from the start of
    // the page's next layout, in which the drawing and its exclusion lie
    // below it and the segment takes the same place again.
    const tableBox = Object.freeze({ sourceIndex: request.input.source.path[0]!, box: bounds });
    const registered = pinned ? null : uprightBlockRegistrationOf(bounds, uprightFragmentOrder(request), [], [
      uprightBlockExclusion(request.input.source, fragmentStartKey, bounds),
    ]);
    const segmentRegistration = !registered ? {} : {
      uprightBlocks: Object.freeze({ ...registered, tableBox }),
      ...(uprightBlockCoversLaterDrawing(context.state, request.input.source, bounds)
        ? {
            uprightPagePin: Object.freeze({
              ...registered, tableBox, segmentPin: Object.freeze({ key: segmentKey, bounds }),
            }),
          }
        : {}),
    };
    return Object.freeze({
      layout,
      blockExtentPt: 0,
      nextCursor,
      ...segmentRegistration,
      placement: Object.freeze({
        coordinateSpace: 'upright-physical' as const,
        xPt: origin.xPt,
        yPt: origin.yPt + layout.flowBounds.yPt,
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
  const columnTopPt = physicalInlineBandTopPt(
    physical, request.location.cursorPt.xPt, request.availableInlineExtentPt,
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
    uprightPhysicalRectToLogical(physical, { xPt, yPt, widthPt, heightPt });
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
