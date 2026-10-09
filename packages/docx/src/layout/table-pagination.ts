import {
  blockNeedsPageOrigin,
  hasPagePlacedContent,
  needsNestedPageOrigins,
  rowNeedsPageOrigins,
  type RetainedTableAcquisition,
} from './table-acquisition.js';
import {
  beginFloatingTablePlacementTransaction,
  floatingTableRegistryDelta,
  resolveFloatingTablePlacementInTransaction,
} from './floating-table-transaction.js';
import {
  ExactConvergenceError,
  convergeExactState,
} from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';
import { paragraphAcquisitionCacheOf } from './runtime-state.js';
import { adjustForWidowOrphan } from '../line-fit-policy.js';
import type { StoryPageFrames } from './story-page-frames.js';
import { sliceParagraphLayout } from './paragraph.js';
import {
  laidOutTableTracks,
  layoutTable,
  measureTableCellBlockFlowHeightPt,
  mergeEndRow,
  projectedMergeRole,
  tablePrefixTracks,
  tableRowBoundaryFootprintsPt,
  type TablePrefixTracks,
} from './table.js';
import {
  wordClipsOverPageCantSplitRow,
  wordDefersCellOwnedAnchorPastPageBand,
  wordRelocatesAuthoredHeightRowAtPageBoundary,
  wordRelocatesParallelParagraphRowCut,
  wordGridFramePreservesEmptyCarrierReference,
} from './table-compatibility.js';
import type {
  FlowBlockPlacement,
  FloatRegistryDeltaPt,
  FloatRegistryEntryPt,
  FloatRegistrySnapshotPt,
  FloatingTablePlacementLayout,
  FloatingTableReferenceFramesPt,
  LayoutServices,
  ParagraphLayout,
  ResolvedFloatingTablePlacementLayout,
  TableCellBlockInput,
  TableCellLayout,
  TableCellLayoutInput,
  TableLayout,
  TableLayoutInput,
  TableRowLayout,
  TableRowLayoutInput,
} from './types.js';

export type TableFragmentOwnership = 'source' | 'repeated-header';

export type BlockContinuationRange =
  | Readonly<{
      kind: 'paragraph';
      blockIndex: number;
      lineStart: number;
      lineEnd: number;
    }>
  | Readonly<{
      kind: 'nested-table';
      blockIndex: number;
      childFragmentIndex: number;
    }>
  | Readonly<{
      kind: 'whole';
      blockIndex: number;
    }>;

export interface TableCellFragmentLayout extends TableCellLayout {
  readonly contentRanges: readonly BlockContinuationRange[];
}

export interface TableRowFragmentLayout extends TableRowLayout {
  readonly logicalRowIndex: number;
  readonly fragmentIndex: number;
  readonly ownership: TableFragmentOwnership;
  readonly occurrenceId: string;
  readonly physicalPageIndex: number;
  readonly displayPageNumber: number;
  readonly cells: readonly TableCellFragmentLayout[];
}

export interface TableFragmentLayout extends TableLayout {
  /** Height beyond the fresh page band retained by an over-page row but
   * clipped from paint. It still occupies physical pages before a following
   * authored page break; see WORD_OVER_PAGE_CELL_BREAK_OCCUPANCY. */
  readonly unpaintedOverflowPt?: number;
  readonly rows: readonly TableRowFragmentLayout[];
  readonly floatingTables: readonly FloatingTablePlacementLayout[];
  readonly resolvedFloatingTables: readonly ResolvedFloatingTablePlacementLayout[];
  readonly resolvedFloatingTableCoordinateSpace?: FloatRegistrySnapshotPt['coordinateSpace'];
}

interface TableCellFragmentCursor {
  readonly blockIndex: number;
  readonly paragraphLineStart: number;
  readonly nestedCursor: TableFragmentCursor | null;
  readonly nestedFragmentIndex: number;
}

export interface TableFragmentCursor {
  readonly rowIndex: number;
  readonly rowFragmentIndex: number;
  readonly cells: readonly TableCellFragmentCursor[];
}

export interface TableFragmentPageContext {
  readonly physicalPageIndex: number;
  readonly displayPageNumber: number;
  readonly occurrenceId: string;
}

export interface PageDependentTableBlockRequest {
  readonly logicalRowIndex: number;
  readonly logicalCellIndex: number;
  readonly sourceBlockIndex: number;
  readonly ownership: TableFragmentOwnership;
  readonly page: TableFragmentPageContext;
  readonly acquired: ParagraphLayout | TableLayout;
  /** Anchor-local point-space exclusions used by final-frame nested floats. */
  readonly floatingTableExclusions?: readonly Readonly<{
    xPt: number;
    yPt: number;
    widthPt: number;
    heightPt: number;
  }>[];
  /** The page translation the paragraph's host flow receives on this page
   * (ParagraphAcquisitionOptions.hostFlowPageTranslationPt), stated by the
   * table's page placement for a cell paragraph whose text boxes hold
   * page-placed content. */
  readonly hostFlowPageTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  readonly paragraphAnchorReferenceDeltaPt?: number;
}

export interface TableFragmentContext {
  readonly availableHeightPt: number;
  readonly freshPageHeightPt: number;
  readonly placement: FlowBlockPlacement;
  readonly services: LayoutServices;
  readonly compatibility: 'word' | 'standard';
  /** Deterministic policy for floating rows taller than a fresh page band. */
  readonly oversizedRowPolicy?: 'split' | 'atomic';
  readonly page: TableFragmentPageContext;
  readonly floatingTableFrames?: Readonly<{
    page: FloatingTableReferenceFramesPt['page'];
    margin: FloatingTableReferenceFramesPt['margin'];
    column: FloatingTableReferenceFramesPt['text'];
  }>;
  readonly floatingTableRegistry?: FloatRegistrySnapshotPt;
  readonly finalPlacementTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  /** Initial ordinary body fragment only; never inherited by nested tables. */
  readonly paragraphAnchorReferenceDeltaPt?: number;
  /** Reacquire only content whose destination page can change its geometry. */
  readonly reacquirePageDependentBlock?: (
    request: PageDependentTableBlockRequest,
  ) => ParagraphLayout | TableLayout;
  /**
   * Page placement of the page-placed content below THIS table's own cells
   * (§17.4.57 positioned tables under in-flow nested tables, text boxes whose
   * stories hold such content): the destination page frames, the table-local
   * to page translation, and re-acquisition of a cell paragraph with its page
   * translation. A nested table's context inherits it only with its own
   * translation (the nested block's page origin stated by the parent's
   * placement), never the parent's.
   */
  readonly pagePlacement?: Readonly<{
    frames: StoryPageFrames;
    translationPt: Readonly<{ xPt: number; yPt: number }>;
    reacquireParagraph: (request: PageDependentTableBlockRequest) => ParagraphLayout | TableLayout;
  }>;
}

export interface TableFragmentResult {
  readonly fragment: TableFragmentLayout | null;
  readonly nextCursor: TableFragmentCursor | null;
  readonly requiresFreshPage: boolean;
  readonly floatingTablePlacements?: readonly ResolvedFloatingTablePlacementLayout[];
  readonly floatingTableRegistryDelta?: FloatRegistryDeltaPt;
}


/** An optional anchor refinement may replace only a pagination-equivalent result. */
export function acceptTableAnchorReferenceRefinement(
  original: TableFragmentResult,
  adjusted: TableFragmentResult,
): TableFragmentResult {
  if (!original.fragment || original.requiresFreshPage) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', 'An anchor refinement requires an admitted table fragment');
  }
  if (adjusted.fragment && !adjusted.requiresFreshPage
    && adjusted.fragment.advancePt === original.fragment.advancePt
    && JSON.stringify(adjusted.fragment.flowBounds) === JSON.stringify(original.fragment.flowBounds)
    && JSON.stringify(adjusted.nextCursor) === JSON.stringify(original.nextCursor)) return adjusted;
  // Keep the complete original transaction, including registry lineage. This
  // does not catch acquisition errors or relax any geometry validation.
  return Object.freeze({
    ...original,
    fragment: Object.freeze({
      ...original.fragment,
      diagnostics: Object.freeze([
        ...(original.fragment.diagnostics ?? []),
        Object.freeze({
          code: 'UNSUPPORTED_FEATURE' as const,
          severity: 'warning' as const,
          source: original.fragment.source,
          message: 'A cell-grid anchor refinement was skipped because it changed table flow; its drawing retains the actual-flow paragraph reference',
        }),
      ]),
    }),
  });
}

interface SelectedCell {
  readonly input: TableCellLayoutInput;
  readonly range: readonly BlockContinuationRange[];
  readonly next: TableCellFragmentCursor;
  readonly complete: boolean;
}

interface SelectedRow {
  readonly input: TableRowLayoutInput;
  readonly logicalRowIndex: number;
  readonly fragmentIndex: number;
  readonly ownership: TableFragmentOwnership;
  readonly ranges: readonly (readonly BlockContinuationRange[])[];
  readonly clipAtPageEnd?: boolean;
  readonly resolvedFloatingTables?: readonly ResolvedFloatingTablePlacementLayout[];
}

const EPSILON_PT = 0.0001;

function emptyCellCursor(): TableCellFragmentCursor {
  return Object.freeze({
    blockIndex: 0,
    paragraphLineStart: 0,
    nestedCursor: null,
    nestedFragmentIndex: 0,
  });
}

export function startTableFragmentCursor(): TableFragmentCursor {
  return Object.freeze({ rowIndex: 0, rowFragmentIndex: 0, cells: Object.freeze([]) });
}

/**
 * Input position of the acquired row a selected occurrence stands for. Rows
 * keep their source logical index; a projected input (a body or story-root
 * cell-owner segment, possibly led by repeated source headers) does not store logical row i at
 * position i, so pagination never treats one as the other.
 */
function inputRowIndexOf(source: RetainedTableAcquisition, logicalRowIndex: number): number {
  if (source.input.rows[logicalRowIndex]?.logicalRowIndex === logicalRowIndex) return logicalRowIndex;
  return source.input.rows.findIndex((row) => row.logicalRowIndex === logicalRowIndex);
}

function sourceRowFor(
  source: RetainedTableAcquisition,
  logicalRowIndex: number,
): TableRowLayoutInput | undefined {
  return source.input.rows[inputRowIndexOf(source, logicalRowIndex)];
}

function leadingHeaderCount(input: TableLayoutInput): number {
  let count = 0;
  while (input.rows[count]?.repeatedHeader === true) count += 1;
  return count;
}

function paginationRowHeight(source: RetainedTableAcquisition, rowIndex: number): number {
  const row = source.layout.rows[rowIndex];
  if (!row) return 0;
  // ECMA-376 §17.4.80 requires exact content to remain inside the authored row
  // box. `layoutTable` has already resolved the compatibility-owned exact track,
  // including its bottom-cell-padding addition; content overflow is
  // paint-time clip ink and must not turn a fitting exact row into continuations.
  if (source.input.rows[rowIndex]?.heightRule === 'exact') return Math.max(0, row.heightPt);
  // A vertically merged owner's content requirement and its physical track can
  // land on different source rows. Considering both permits a legal boundary
  // through the span without rewriting restart/continue semantics.
  return Math.max(0, row.heightPt, row.contentHeightPt);
}

function nestedTableFragmentContext(
  source: RetainedTableAcquisition,
  parentCell: TableCellLayoutInput,
  context: TableFragmentContext,
  availableHeightPt: number,
  nestedBlock?: TableCellBlockInput,
): TableFragmentContext {
  const retainedCell = source.layout.rows
    .flatMap((row) => row.cells)
    .find((cell) => cell.id === parentCell.id);
  if (!retainedCell) {
    throw new LayoutInvariantError(
      'INVALID_REFERENCE',
      `nested table fragment lost parent cell geometry: ${parentCell.id}`,
    );
  }
  const bounds = Object.freeze({
    xPt: 0,
    yPt: 0,
    widthPt: retainedCell.contentBounds.widthPt,
    heightPt: Math.max(0, availableHeightPt),
  });
  // The parent's page translation is not this table's: the nested block's own
  // page origin, stated by the parent's page placement, replaces it.
  const { pagePlacement: parentPagePlacement, ...inherited } = context;
  const origin = nestedBlock ? nestedPageOrigins.get(nestedBlock) : undefined;
  return Object.freeze({
    ...inherited,
    availableHeightPt: bounds.heightPt,
    placement: Object.freeze({
      ...context.placement,
      container: Object.freeze({ ...context.placement.container, bounds }),
      cursor: Object.freeze({ xPt: 0, yPt: 0 }),
      availableBounds: bounds,
    }),
    // With its page origin known, the fragment's own positioned children are
    // resolved through it, not through the parent's translation.
    ...(origin && parentPagePlacement ? {
      pagePlacement: Object.freeze({ ...parentPagePlacement, translationPt: origin }),
      finalPlacementTranslationPt: origin,
    } : {}),
  });
}

function paginationRowHeightForOccurrence(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  rowIndex: number,
  context: TableFragmentContext,
): number {
  if (row === source.input.rows[rowIndex]) return paginationRowHeight(source, rowIndex);
  const occurrence = layoutTable({
    ...source.input,
    id: `${source.input.id}:row-occurrence:${context.page.occurrenceId}:${row.logicalRowIndex}`,
    rows: [row],
  }, context.placement, context.services).layout;
  return Math.max(0, occurrence.rows[0]?.heightPt ?? occurrence.advancePt);
}

function paginationRowTrackHeightForOccurrence(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  rowIndex: number,
  context: TableFragmentContext,
): number {
  if (row === source.input.rows[rowIndex]) {
    return Math.max(0, source.layout.rows[rowIndex]?.heightPt ?? 0);
  }
  return paginationRowHeightForOccurrence(source, row, rowIndex, context);
}

/**
 * The last source row that can influence the completed partial row's physical
 * track. Only rows inside the transitive closure of the vertical-merge
 * intervals overlapping the completed row matter: the deficit relocation in
 * table.ts resolveRowHeights reads and grows heights strictly inside each
 * owner's interval, so intervals disjoint from the completed row cannot reach
 * its track. The window is closed transitively because an interval opening
 * inside it may itself reach further down.
 *
 * An interval opens where a cell's role in the window's own grid (`row`
 * first, then the source rows below it) is a restart: the authored one, or
 * the projected empty owner of a cell-owner segment's opening row
 * (`segmentOpeningLogicalRowIndex`, table.ts projectedMergeRole), which every
 * fragment of that row opens. Reading the authored role alone would close the
 * projected owner's interval at the window's end, and its margin deficit
 * would then land on the completed row instead of the interval's last
 * growable track.
 */
export function completedPartialRowWindowEnd(
  row: TableRowLayoutInput,
  sourceRows: readonly TableRowLayoutInput[],
  rowIndex: number,
  segmentOpeningLogicalRowIndex?: TableLayoutInput['segmentOpeningLogicalRowIndex'],
): number {
  let windowEnd = rowIndex;
  for (let scan = rowIndex; scan <= windowEnd && scan < sourceRows.length; scan += 1) {
    const scanned = scan === rowIndex ? row : sourceRows[scan]!;
    const above = scan === rowIndex ? undefined : scan === rowIndex + 1 ? row : sourceRows[scan - 1];
    for (const cell of scanned.cells) {
      if (projectedMergeRole(segmentOpeningLogicalRowIndex, above, scanned, cell) !== 'restart') continue;
      windowEnd = Math.max(
        windowEnd,
        mergeEndRow(sourceRows, scan, cell.columnStart, cell.columnSpan),
      );
    }
  }
  return windowEnd;
}

function completedPartialRowTrackHeight(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  rowIndex: number,
  context: TableFragmentContext,
): number {
  // A merged owner's deficit is assigned across its complete continuation
  // interval. Resolve that interval once so the completed partial row is
  // charged its physical track, not the owner's full isolated content height.
  // Only the merge window above can influence that track, plus the row right
  // after it, whose top borders and cell spacing resolve the window's bottom
  // boundary. Lay out that bounded window instead of the whole remaining
  // table: with no merge it is two rows, so pagination no longer costs
  // O(remaining rows) per completed partial row.
  const sourceRows = source.input.rows;
  const windowEnd = completedPartialRowWindowEnd(
    row, sourceRows, rowIndex, source.input.segmentOpeningLogicalRowIndex,
  );
  const sliceEnd = Math.min(sourceRows.length, windowEnd + 2);
  const occurrence = layoutTable({
    ...source.input,
    id: `${source.input.id}:completed-partial:${context.page.occurrenceId}:${row.logicalRowIndex}`,
    rows: [row, ...sourceRows.slice(rowIndex + 1, sliceEnd)],
  }, context.placement, context.services).layout;
  return Math.max(0, occurrence.rows[0]?.heightPt ?? 0);
}

function rowRanges(row: TableRowLayoutInput): readonly (readonly BlockContinuationRange[])[] {
  return row.cells.map((cell) => cell.blocks.map((block) => ({
    kind: 'whole' as const,
    blockIndex: block.sourceBlockIndex,
  })));
}

function rowForOccurrence(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  context: TableFragmentContext,
): TableRowLayoutInput {
  const reacquire = context.reacquirePageDependentBlock;
  if (!reacquire || !row.cells.some((cell) => (
    cell.blocks.some((block) => block.pageDependent === true)
  ))) return row;
  return {
    ...row,
    cells: row.cells.map((cell, logicalCellIndex) => ({
      ...cell,
      blocks: cell.blocks.map((block) => block.pageDependent === true
        ? {
            ...block,
            layout: reacquire({
              logicalRowIndex: row.logicalRowIndex,
              logicalCellIndex,
              sourceBlockIndex: block.sourceBlockIndex,
              ownership,
              page: context.page,
              acquired: block.layout,
            }),
          }
        : block),
    })),
  };
}

type PagePoint = Readonly<{ xPt: number; yPt: number }>;

/**
 * Page position of a nested table block's local origin in the occurrence that
 * holds it, stated by that occurrence's page placement. A fragment of the
 * nested table taken there places its own page-placed content through it, at
 * every nesting depth. Keyed by the immutable placed block input.
 */
const nestedPageOrigins = new WeakMap<TableCellBlockInput, PagePoint>();

/**
 * A retained table's §17.4.57 positioned children by host cell (indices into
 * `floatingTables`, in source order), and its acquired cells by id. Both are
 * fixed by the immutable acquisition, so each table is indexed once, and a
 * row finds its own children in time proportional to its cells and children
 * rather than to every child of the table. Keyed by the acquisition, released
 * with it.
 */
type FloatingTableIndex = Readonly<{
  byHostCell: ReadonlyMap<string, readonly number[]>;
  cellById: ReadonlyMap<string, TableLayout['rows'][number]['cells'][number]>;
}>;

const floatingTableIndexes = new WeakMap<RetainedTableAcquisition, FloatingTableIndex>();

function floatingTableIndexOf(source: RetainedTableAcquisition): FloatingTableIndex {
  const known = floatingTableIndexes.get(source);
  if (known) return known;
  const byHostCell = new Map<string, number[]>();
  source.floatingTables.forEach((occurrence, index) => {
    const hosted = byHostCell.get(occurrence.hostCellId);
    if (hosted) hosted.push(index);
    else byHostCell.set(occurrence.hostCellId, [index]);
  });
  const cellById = new Map<string, TableLayout['rows'][number]['cells'][number]>();
  for (const row of source.layout.rows) {
    for (const cell of row.cells) if (!cellById.has(cell.id)) cellById.set(cell.id, cell);
  }
  const index: FloatingTableIndex = Object.freeze({ byHostCell, cellById });
  floatingTableIndexes.set(source, index);
  return index;
}

/**
 * Placement probes of one fragment's growing list of selected rows. A probe
 * of a candidate is the row the fragment lays out after those rows — the
 * last row of the table laid out from them and the candidate — taken from the
 * table's exact prefix tracks (table.ts tablePrefixTracks). Those resolve the
 * rows before the list's last one once each and resolve a probe in time
 * bounded by the row's own cells and the merges crossing it, so one list's
 * probes cost work proportional to its rows whatever the merges: under merges
 * continuing through every row, for `exact` rows (whose merge deficit lands
 * above them) and for full-row probes (cell heights, vAlign offsets, merge
 * growth on the candidate) alike. Released with the list. A table laid out
 * whole is not probed: it is placed against its own final layout
 * ({@link layoutNestedWholeAt}).
 */
type FragmentRowTracks = TablePrefixTracks;

function startFragmentRowTracks(source: RetainedTableAcquisition): FragmentRowTracks {
  return tablePrefixTracks(source.input);
}

/** The candidate as the last row of the whole prefix: the probe of a grid the
 * prefix tracks do not resolve (two owners claiming one column of a row).
 *
 * Reachability: canonical DOCX acquisition does not produce such a grid. It
 * assigns each row's cells sequential grid columns from the row's gridBefore,
 * each span at least one column and clamped to the grid
 * (table-acquisition.ts, table-source-acquisition.ts), so the cells of one
 * row claim disjoint columns; and a merge stays open into a row only through
 * a continuation cell of that row with the owner's own start and span
 * (table.ts resolveRowTrack, continuesMerge; ECMA-376 §17.4.84 merges cells
 * of the same grid columns), so every open owner's columns are those of one
 * of the row's own cells. The public document model carries no column start
 * to break this. Only a TableLayoutInput built by hand outside acquisition
 * reaches here.
 *
 * Cost: laying out the prefix is proportional to it, so a fragment probing
 * such a grid row by row costs the square of its rows. Each probe is charged
 * to the session's acquisition budget (runtime-state.ts), as a speculative
 * whole layout nested in another is: the fallback is bounded in count, not
 * made linear in the input, and needs no budget of its own since canonical
 * input never takes it. */
function wholePrefixProbe(
  source: RetainedTableAcquisition,
  preceding: readonly TableRowLayoutInput[],
  candidate: TableRowLayoutInput,
  context: TableFragmentContext,
): TableRowLayout {
  paragraphAcquisitionCacheOf(context.services)?.noteMiss();
  const laidOut = layoutTable({
    ...source.input,
    id: `${source.input.id}:page-origin-probe:${context.page.occurrenceId}:${candidate.logicalRowIndex}`,
    rows: [...preceding, candidate],
  }, context.placement, context.services).layout;
  const row = laidOut.rows[laidOut.rows.length - 1];
  if (!row) throw new Error('Page origin probe lost its row');
  return row;
}

/** The candidate laid out after `preceding` as the fragment will lay it out,
 * in fragment coordinates. */
function probeFragmentRow(
  source: RetainedTableAcquisition,
  preceding: readonly TableRowLayoutInput[],
  candidate: TableRowLayoutInput,
  context: TableFragmentContext,
  tracks: FragmentRowTracks,
): TableRowLayout {
  return tracks.row(preceding, candidate, context.placement)
    ?? wholePrefixProbe(source, preceding, candidate, context);
}

/**
 * The placement a nested table is acquired with (table-acquisition.ts) in a
 * cell content box `widthPt` wide, and the page placement of the content
 * below its own cells through `origin`, the page position of its local
 * origin.
 */
function nestedWholeContext(
  context: TableFragmentContext,
  nested: RetainedTableAcquisition,
  widthPt: number,
  origin: PagePoint,
): TableFragmentContext {
  const bounds = Object.freeze({ xPt: 0, yPt: 0, widthPt, heightPt: 1 });
  return Object.freeze({
    ...context,
    paragraphAnchorReferenceDeltaPt: undefined,
    placement: Object.freeze({
      container: Object.freeze({ id: nested.input.flowDomainId, kind: 'tableCell' as const, bounds }),
      cursor: Object.freeze({ xPt: 0, yPt: 0 }),
      availableBounds: bounds,
    }),
    pagePlacement: Object.freeze({ ...context.pagePlacement!, translationPt: origin }),
  });
}

/** Whether a table laid out whole has anything only a page position places:
 * §17.4.57 positioned tables at any depth, or text boxes holding them. */
export function hasPagePlacedTableContent(source: RetainedTableAcquisition): boolean {
  return hasPagePlacedContent(source);
}

/** A story table laid out whole at a known page translation (the context's
 * `pagePlacement.translationPt`), as {@link layoutNestedWhole} lays out a
 * nested one: positioned children given final frames, page-placed content
 * below its cells placed through it. */
export function layoutWholeTableOnPage(
  source: RetainedTableAcquisition,
  context: TableFragmentContext,
): TableLayout {
  return layoutNestedWhole(source, context);
}

/**
 * A nested table laid out whole at the page origin its context's placement
 * states (an in-flow nested table in a whole row, a positioned table at its
 * final frame): the content below its cells is placed through it
 * ({@link rowPagePlacement}) and its own §17.4.57 positioned children are
 * given their final frames there, row by row in source order, by the same
 * {@link finalFrameRow} step a fragment holding the whole table takes, so
 * their anchor paragraphs wrap and their own content is placed through them
 * in turn. Both read each row where the whole table's own layout puts it
 * ({@link layoutNestedWholeAt}). The resolutions are this layout's own page
 * placements; like any final frame they avoid the floats of the context's
 * registry snapshot, and they are not committed to the page registry, which
 * only a paginated table's direct children enter.
 */
function layoutNestedWhole(
  nested: RetainedTableAcquisition,
  context: TableFragmentContext,
): TableLayout {
  // Each pass of a solve in such a layout that moves an origin repeats every
  // layout nested in it, so nesting multiplies the per-loop guards (a pass
  // revisiting an origin reuses its layout: rowPagePlacement,
  // resolveFinalFrameChild). A layout inside another one is therefore
  // charged to the session's acquisition budget (runtime-state.ts), which
  // bounds the total; the outermost one is bounded by its caller's own
  // loops, per row and pass.
  if (nestedWholeDepth > 0) {
    paragraphAcquisitionCacheOf(context.services)?.noteMiss();
  }
  nestedWholeDepth += 1;
  try {
    return layoutNestedWholeAt(nested, context);
  } finally {
    nestedWholeDepth -= 1;
  }
}

/** Depth of {@link layoutNestedWhole} calls in progress (layout is synchronous). */
let nestedWholeDepth = 0;

/**
 * The whole table's rows placed where its own final layout puts them (library
 * policy). A row of a table laid out whole is not the last row of the rows
 * above it: a merge continuing below it gives its deficit to a later row (or,
 * through `exact` rows, to an earlier one; table.ts resolveRowTrack), so its
 * track — and with it every centered or bottom vAlign offset and every later
 * row's top — is the whole table's, not its prefix's. Each row's page-placed
 * content (nested tables through their page origin, text-box paragraphs
 * through their page translation; {@link rowPagePlacement}) and the track a
 * row hosting positioned children resolves them in (its top, height and
 * owned merges' extent, so its anchors and vAlign offsets; finalFrameRow,
 * table.ts laidOutTableTracks) are therefore read from the table laid out
 * from every row. That content can change the rows it is placed in (an
 * anchor paragraph wrapping below a child, a nested table growing), and so
 * the layout it is read from, so the pair is solved to an exact fixed point:
 * a pass reads the placements and tracks from the previous pass's layout
 * (the first from the acquired rows), prepares every row from them in source
 * order — each positioned child resolved once per pass, against the registry
 * of the children before it, starting from its previous pass's resolution —
 * and lays the prepared rows out; the layout whose placements and tracks are
 * the ones its rows were prepared with is accepted. No threshold is applied;
 * a cycle or exhaustion fails closed, like the per-row loops.
 *
 * Cost: a pass lays the table out once, builds its tracks once when a row
 * hosts a child, and reads each row once, so a pass is proportional to the
 * table, and the passes are bounded by the resource guard. A pass that keeps
 * a row's placement reuses its placed row, and one that also keeps its track
 * and the registry before it reuses its prepared row, so only rows whose
 * geometry moved are prepared again; nested layouts placed through an origin
 * a pass revisits are reused across passes. The confirming pass prepares and
 * lays out nothing.
 */
function layoutNestedWholeAt(
  nested: RetainedTableAcquisition,
  context: TableFragmentContext,
): TableLayout {
  const sourceRows = nested.input.rows;
  const layoutRows = (rows: readonly TableRowLayoutInput[]) => layoutTable(
    { ...nested.input, rows }, context.placement, context.services,
  ).layout;
  const placedWhole: PlacedWholeLayouts = new Map();
  const placements: readonly (RowPagePlacement | null)[] = needsNestedPageOrigins(nested)
    ? sourceRows.map((row) => rowPagePlacement(nested, row, 'source', context, placedWhole))
    : [];
  const origin = context.pagePlacement?.translationPt;
  const resolvesChildren = nested.floatingTables.length > 0 && origin !== undefined
    && context.floatingTableFrames !== undefined && context.reacquirePageDependentBlock !== undefined;
  if (!resolvesChildren && !placements.some((placement) => placement !== null)) {
    return layoutRows(sourceRows);
  }
  const wholeContext: TableFragmentContext = resolvesChildren && origin
    ? Object.freeze({ ...context, finalPlacementTranslationPt: origin })
    : context;
  const initialRegistry: readonly FloatRegistryEntryPt[] = context.floatingTableRegistry?.entries ?? [];
  const initialParagraphId = context.floatingTableRegistry?.nextParagraphId ?? 0;
  // Only a row hosting a positioned child reads its track (finalFrameRow).
  const { byHostCell } = floatingTableIndexOf(nested);
  const hostsChild = sourceRows.map((row) => (
    resolvesChildren && row.cells.some((cell) => byHostCell.has(cell.id))
  ));
  type PreparedRow = Readonly<{
    placementKey: string;
    placed: TableRowLayoutInput;
    trackKey: string | null;
    registry: readonly FloatRegistryEntryPt[];
    nextParagraphId: number;
    prepared: ReturnType<typeof finalFrameRow>;
  }>;
  type Pass = Readonly<{
    state: string;
    rows: readonly TableRowLayoutInput[];
    prepared: readonly PreparedRow[];
    resolved: readonly ResolvedFloatingTablePlacementLayout[];
    layout: TableLayout;
  }>;
  const step = (previous: Pass | null): Pass => {
    const laidOut = previous?.layout ?? layoutRows(sourceRows);
    const laidOutRows = previous?.rows ?? sourceRows;
    const laidOutRow = (rowIndex: number): TableRowLayout => {
      const row = laidOut.rows[rowIndex];
      if (!row) throw new LayoutInvariantError('INVALID_REFERENCE', `whole table lost row ${rowIndex}`);
      return row;
    };
    const placementStates = sourceRows.map((_row, rowIndex) => (
      placements[rowIndex]?.stateAt(laidOutRows[rowIndex]!, laidOutRow(rowIndex)) ?? null
    ));
    // A hosting row's track in this layout: its top, its height and, for a
    // merge it owns, the merge's last row are the whole table's, so its
    // children's anchors and its cells' vAlign offsets are the ones this
    // layout paints (table.ts laidOutTableTracks), not those of the row laid
    // out by itself, where a merge it starts would end in it.
    const trackKeys = sourceRows.map((_row, rowIndex) => (
      hostsChild[rowIndex] ? rowTrackKey(laidOutRow(rowIndex)) : null
    ));
    const state = JSON.stringify([placementStates, trackKeys]);
    // The rows were prepared with this layout's own placements and tracks.
    if (previous?.state === state) return previous;
    let tracks: ReturnType<typeof laidOutTableTracks> | null = null;
    const trackOf = (rowIndex: number): RowTrack => (candidate) => {
      const built = tracks ?? laidOutTableTracks({ ...nested.input, rows: laidOutRows }, laidOut, context.placement);
      tracks = built;
      return { input: candidate, row: built.row(rowIndex, candidate) };
    };
    let registry = initialRegistry;
    let nextParagraphId = initialParagraphId;
    const rows: TableRowLayoutInput[] = [];
    const preparedRows: PreparedRow[] = [];
    const resolved: ResolvedFloatingTablePlacementLayout[] = [];
    sourceRows.forEach((sourceRow, rowIndex) => {
      const placementState = placementStates[rowIndex] ?? null;
      const placementKey = JSON.stringify(placementState);
      const trackKey = trackKeys[rowIndex] ?? null;
      const known = previous?.prepared[rowIndex];
      const placed = placementState === null
        ? sourceRow
        : known?.placementKey === placementKey
          ? known.placed
          : placements[rowIndex]!.rowFor(placementState);
      // A row whose track moved starts from its previous pass's children
      // (finalFrameRow seed).
      const prepared = trackKey === null
        ? { row: placed, resolved: [], registry, nextParagraphId }
        : known && known.placed === placed && known.trackKey === trackKey
          && known.registry === registry && known.nextParagraphId === nextParagraphId
          ? known.prepared
          : finalFrameRow(
            nested, placed, 'source', trackOf(rowIndex), wholeContext, registry, nextParagraphId,
            startTableFragmentCursor(), () => true, known?.prepared.resolved,
          );
      preparedRows.push({ placementKey, placed, trackKey, registry, nextParagraphId, prepared });
      rows.push(prepared.row);
      resolved.push(...prepared.resolved);
      registry = prepared.registry;
      nextParagraphId = prepared.nextParagraphId;
    });
    return Object.freeze({
      state,
      rows: Object.freeze(rows),
      prepared: Object.freeze(preparedRows),
      resolved: Object.freeze(resolved),
      layout: layoutRows(rows),
    });
  };
  let solved: Pass;
  try {
    solved = convergeExactState<Pass>({
      step,
      stateOf: (pass) => pass.state,
      // Resource guard only; exact equality determines acceptance.
      limit: 16,
    }).value;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `whole table page placement did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
  const { layout, resolved } = solved;
  return resolved.length === 0 ? layout : Object.freeze({
    ...layout,
    resolvedFloatingTables: resolved,
    ...(context.floatingTableRegistry ? {
      resolvedFloatingTableCoordinateSpace: context.floatingTableRegistry.coordinateSpace,
    } : {}),
  });
}

/** The page position of the content below one cell of a row. */
type CellPagePlacement = Readonly<{
  nested: readonly Readonly<{ blockIndex: number; origin: PagePoint; widthPt: number }>[];
  /** Paragraphs whose text boxes hold page-placed content: their page
   * translation. */
  hostFlows: readonly Readonly<{ blockIndex: number; translation: PagePoint }>[];
}>;

const UNPLACED_CELL: CellPagePlacement = Object.freeze({ nested: [], hostFlows: [] });

/** Nested tables laid out whole through a page origin, by table, origin and
 * width: one placement's (every pass of its solve reuses an origin it
 * revisits), released with it. */
type PlacedWholeLayouts = Map<string, TableLayout>;

/**
 * Page placement of the content below a row's cells (library policy): a
 * nested table block holding page-placed content (§17.4.57 positioned tables
 * at any depth below it) and a cell paragraph whose text boxes hold such
 * content are given their page position. `stateAt` reads that position from
 * a laid-out occurrence of the row; `rowFor` places the content there. A
 * nested table laid out whole here is laid out whole through its origin
 * ({@link layoutNestedWhole}); a fragment of it is taken through the origin.
 * A paragraph is re-acquired with its page translation, so its text boxes'
 * stories are given their page frames. Without a page placement (a step that
 * owns no page) the content keeps its acquisition (null), as a table's direct
 * positioned children do without frames.
 */
type RowPagePlacement = Readonly<{
  /** The placement where `laidOutRow`, the occurrence `laidOutInput` laid
   * out (for a continued row, its remainder), puts the row's content. */
  stateAt: (laidOutInput: TableRowLayoutInput, laidOutRow: TableRowLayout) => readonly CellPagePlacement[];
  rowFor: (state: readonly CellPagePlacement[]) => TableRowLayoutInput;
}>;

function rowPagePlacement(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  context: TableFragmentContext,
  placedWholeLayouts: PlacedWholeLayouts,
  cursor?: TableFragmentCursor,
): RowPagePlacement | null {
  const keepsInsertionReference = (cellIndex: number, blockIndex: number): boolean => (
    context.paragraphAnchorReferenceDeltaPt !== undefined
    && ownership === 'source'
    && (!cursor || (cursor.rowIndex === 0 && cursor.rowFragmentIndex === 0 && cursor.cells.length === 0))
    && row.logicalRowIndex === 0 && !row.repeatedHeader
    && cellIndex === 0 && blockIndex === 0
    && wordGridFramePreservesEmptyCarrierReference(row.cells[cellIndex]!)
  );
  if (!rowNeedsPageOrigins(source, row) && !keepsInsertionReference(0, 0)) return null;
  const owner = context.pagePlacement;
  if (!owner) return null;
  const translation = owner.translationPt;
  // Blocks before a continued row's cell cursor were placed by earlier
  // fragments.
  const continuing = cursor !== undefined && cursor.rowFragmentIndex > 0;
  const firstBlock = (cellIndex: number) => (
    continuing ? cursor.cells[cellIndex]?.blockIndex ?? 0 : 0
  );
  const stateAt = (
    candidate: TableRowLayoutInput,
    laidOutRow: TableRowLayout,
  ): readonly CellPagePlacement[] => (
    candidate.cells.map((cell, cellIndex): CellPagePlacement => {
      const laidOutCell = laidOutRow.cells[cellIndex];
      const start = firstBlock(cellIndex);
      if (!laidOutCell || !cell.blocks.slice(start).some((block) => (
        blockNeedsPageOrigin(source, block) || keepsInsertionReference(cellIndex, start)
      ))) return UNPLACED_CELL;
      const placedAt = (blockIndex: number) => laidOutCell.blocks[blockIndex - start];
      const cellTopPt = translation.yPt + laidOutCell.flowBounds.yPt;
      const contentXPt = translation.xPt + laidOutCell.contentBounds.xPt;
      const nested: { blockIndex: number; origin: PagePoint; widthPt: number }[] = [];
      const hostFlows: { blockIndex: number; translation: PagePoint }[] = [];
      for (let blockIndex = start; blockIndex < cell.blocks.length; blockIndex += 1) {
        const block = cell.blocks[blockIndex]!;
        const placed = placedAt(blockIndex);
        if (!placed || (!blockNeedsPageOrigin(source, block)
          && !keepsInsertionReference(cellIndex, blockIndex))) continue;
        const placedTopPt = cellTopPt + placed.offsetPt;
        if (block.layout.kind === 'table') {
          // A nested table's local origin is placed at its cell content x and
          // block offset (table.ts placedChildInkBounds).
          nested.push({
            blockIndex,
            origin: { xPt: contentXPt, yPt: placedTopPt },
            widthPt: laidOutCell.contentBounds.widthPt,
          });
        } else if (block.layout.kind === 'paragraph') {
          // Page point → this paragraph's acquired coordinates (the inverse
          // of table.ts child placement): the translation its host flow, and
          // so its text boxes' stories, receive on the page.
          hostFlows.push({
            blockIndex,
            translation: {
              xPt: contentXPt - block.layout.flowBounds.xPt,
              yPt: placedTopPt - block.layout.flowBounds.yPt,
            },
          });
        }
      }
      return { nested, hostFlows };
    })
  );
  const placedWhole = (
    block: TableCellBlockInput,
    at: Readonly<{ origin: PagePoint; widthPt: number }>,
  ): TableLayout => {
    const key = JSON.stringify([block.layout.id, at.origin, at.widthPt]);
    const known = placedWholeLayouts.get(key);
    if (known) return known;
    const nested = source.nestedById[block.layout.id]!;
    const layout = layoutNestedWhole(
      nested,
      nestedWholeContext(context, nested, at.widthPt, at.origin),
    );
    placedWholeLayouts.set(key, layout);
    return layout;
  };
  const rowFor = (state: readonly CellPagePlacement[]): TableRowLayoutInput => ({
    ...row,
    cells: row.cells.map((cell, logicalCellIndex) => {
      const cellState = state[logicalCellIndex];
      if (!cellState || cellState === UNPLACED_CELL) return cell;
      return {
        ...cell,
        blocks: cell.blocks.map((block, blockIndex): TableCellBlockInput => {
          const nestedAt = cellState.nested.find((candidate) => candidate.blockIndex === blockIndex);
          if (nestedAt) {
            // A nested table this fragment continues keeps its acquisition;
            // its remainder is taken through the origin.
            const continues = continuing
              && blockIndex === firstBlock(logicalCellIndex)
              && cursor.cells[logicalCellIndex]?.nestedCursor != null;
            const placed: TableCellBlockInput = continues
              ? { ...block }
              : { ...block, layout: placedWhole(block, nestedAt) };
            nestedPageOrigins.set(placed, nestedAt.origin);
            return placed;
          }
          const hostFlow = cellState.hostFlows
            .find((candidate) => candidate.blockIndex === blockIndex)?.translation;
          if (!hostFlow) return block;
          return {
            ...block,
            layout: owner.reacquireParagraph({
              logicalRowIndex: row.logicalRowIndex,
              logicalCellIndex,
              sourceBlockIndex: block.sourceBlockIndex,
              ownership,
              page: context.page,
              acquired: block.layout,
              hostFlowPageTranslationPt: hostFlow,
              ...(keepsInsertionReference(logicalCellIndex, blockIndex) ? {
                paragraphAnchorReferenceDeltaPt: context.paragraphAnchorReferenceDeltaPt,
              } : {}),
            }),
          };
        }),
      };
    }),
  });
  return Object.freeze({ stateAt, rowFor });
}

/**
 * {@link rowPagePlacement} of a row of a paginated fragment. The row (or, for
 * a continued row, the remainder this fragment lays out; for a row this
 * fragment cuts, the cut occurrence) is probed after the rows already
 * selected in this fragment, as the fragment lays them out
 * ({@link probeFragmentRow}), or, once the fragment has been materialized,
 * in that materialization's own track of the row (`frame.final`; see
 * {@link takeTableFragment}). Everything is repeated until exact (re-acquired
 * content can move the blocks after it). `frame.seed`, the row this placed in
 * a previous pass, only starts the iteration; `frame.used` is told the track
 * the accepted placement was read from.
 */
function placeRowNestedContent(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  precedingRows: () => readonly TableRowLayoutInput[],
  ownership: TableFragmentOwnership,
  context: TableFragmentContext,
  tracks: FragmentRowTracks,
  cursor?: TableFragmentCursor,
  cutOf?: (candidate: TableRowLayoutInput) => TableRowLayoutInput | null,
  frame?: FragmentRowFrame,
): TableRowLayoutInput {
  const placement = rowPagePlacement(source, row, ownership, context, new Map(), cursor);
  if (!placement) return row;
  const preceding = precedingRows();
  // A continued row is probed as the remainder this fragment lays out.
  const continuing = cursor !== undefined && cursor.rowFragmentIndex > 0;
  const stateOf = (candidate: TableRowLayoutInput): readonly CellPagePlacement[] => {
    // A row this fragment cuts is probed as the cut occurrence: its track and
    // so every non-top vAlign offset are the cut's, not the uncut row's.
    const probed = cutOf?.(candidate) ?? (continuing
      ? remainingRowAtCursor(source, candidate, cursor, context, new Map())
      : candidate);
    const laidOut = frame?.final
      ? frame.final(probed)
      : probeFragmentRow(source, preceding, probed, context, tracks);
    frame?.used(laidOut);
    return placement.stateAt(candidate, laidOut);
  };
  try {
    return convergeExactState<Readonly<{ row: TableRowLayoutInput; state: string }>>({
      seedState: JSON.stringify(null),
      step: (previous) => {
        const state = stateOf(previous?.row ?? frame?.seed ?? row);
        return Object.freeze({ row: placement.rowFor(state), state: JSON.stringify(state) });
      },
      stateOf: (pass) => pass.state,
      // Resource guard only; exact equality determines acceptance.
      limit: 16,
    }).value.row;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `page placement below a table row did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
}

/**
 * A §17.4.57 positioned child with page-placed content of its own (library
 * policy). The positioned nested table is placed by the final frame resolved
 * here, so the page position of its local origin is that frame's: the
 * content below its cells (its own positioned children, at any in-flow depth
 * below it) is placed through it, exactly as a paginated table places it
 * through its page translation. The frame depends on the placed child's
 * extent (the vertical page clamp, the registry avoidance) and the placement
 * on the frame, so the pair is iterated to an exact fixed point: the frame
 * the placed child resolves to is the frame it was placed through. No
 * threshold is applied; a cycle or exhaustion fails closed like the
 * final-frame reflow around it. Without a page placement the child keeps its
 * acquisition.
 */
function resolveFinalFrameChild<R extends Readonly<{ placement: ResolvedFloatingTablePlacementLayout }>>(
  source: RetainedTableAcquisition,
  occurrence: RetainedTableAcquisition['floatingTables'][number],
  context: TableFragmentContext,
  resolveWith: (child: TableLayout) => R,
): R {
  // The placement being resolved already found this child.
  const nested = source.nestedById[occurrence.tableId]!;
  if (!hasPagePlacedContent(nested) || !context.pagePlacement) return resolveWith(nested.layout);
  const hostCell = floatingTableIndexOf(source).cellById.get(occurrence.hostCellId);
  if (!hostCell) throw new LayoutInvariantError('INVALID_REFERENCE', `${occurrence.hostCellId} is not a cell`);
  // The child is acquired in its host cell's content box (table-acquisition.ts).
  const widthPt = hostCell.contentBounds.widthPt;
  // The child placed through each origin this solve visits, laid out once: a
  // pass that confirms the previous origin (the fixed point) is the previous
  // pass's child, so it is not laid out — with every level nested in it —
  // again. Released with the solve.
  const children = new Map<string, TableLayout>();
  const childAt = (origin: PagePoint): TableLayout => {
    const key = JSON.stringify(origin);
    const known = children.get(key);
    if (known) return known;
    const child = layoutNestedWhole(nested, nestedWholeContext(context, nested, widthPt, origin));
    children.set(key, child);
    return child;
  };
  const originOf = (resolved: R, child: TableLayout): PagePoint => Object.freeze({
    xPt: resolved.placement.xPt - child.flowBounds.xPt,
    yPt: resolved.placement.yPt - child.flowBounds.yPt,
  });
  type Pass = Readonly<{ resolution: R; origin: PagePoint }>;
  try {
    return convergeExactState<Pass>({
      step: (previous) => {
        const origin = previous?.origin ?? originOf(resolveWith(nested.layout), nested.layout);
        const child = childAt(origin);
        const resolution = resolveWith(child);
        return Object.freeze({ resolution, origin: originOf(resolution, child) });
      },
      stateOf: (pass) => JSON.stringify(pass.origin),
      // Resource guard only; exact equality determines acceptance.
      limit: 16,
    }).value.resolution;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `positioned child page placement did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
}

function remainingRowAtCursor(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  cursor: TableFragmentCursor,
  context: TableFragmentContext,
  requiredAnchorByCell: ReadonlyMap<string, number>,
): TableRowLayoutInput {
  return {
    ...row,
    heightPt: null,
    heightRule: 'auto',
    cells: row.cells.map((cell, cellIndex) => {
      const cellCursor = cursor.cells[cellIndex] ?? emptyCellCursor();
      return {
        ...cell,
        blocks: cell.blocks.slice(cellCursor.blockIndex).map((block, blockOffset) => {
          if (blockOffset === 0 && cellCursor.nestedCursor && block.layout.kind === 'table') {
            const nested = source.nestedById[block.layout.id];
            if (nested) {
              const remaining = takeTableFragment(
                nested,
                cellCursor.nestedCursor,
                nestedTableFragmentContext(
                  source,
                  cell,
                  context,
                  context.freshPageHeightPt,
                  block,
                ),
              );
              const requiredAnchor = requiredAnchorByCell.get(cell.id);
              if (remaining.nextCursor && requiredAnchor !== undefined
                && block.sourceBlockIndex < requiredAnchor) {
                throw new Error(
                  'Floating table anchor cannot follow an incomplete nested-table candidate',
                );
              }
              if (remaining.fragment) return { ...block, layout: remaining.fragment };
            }
          }
          if (blockOffset !== 0
            || cellCursor.paragraphLineStart === 0
            || block.layout.kind !== 'paragraph') return block;
          return {
            ...block,
            layout: paragraphSlice(
              block.layout,
              cellCursor.paragraphLineStart,
              block.layout.lines.length,
            ),
          };
        }),
      };
    }),
  };
}

/**
 * Where a row being prepared is laid out: `candidate` (the row as prepared so
 * far; `remaining`, its remainder at the preparing fragment's cursor) as the
 * layout holding it lays it out, and the input that layout is given for it.
 */
type RowTrack = (
  candidate: TableRowLayoutInput,
  remaining: TableRowLayoutInput,
) => Readonly<{ input: TableRowLayoutInput; row: TableRowLayout }>;

/** The remainder laid out by itself `offsetPt` below the table's top: a
 * first estimate of a fragment row's track, before any materialization of
 * the fragment is known ({@link takeTableFragment}). */
function aloneRowTrack(
  source: RetainedTableAcquisition,
  context: TableFragmentContext,
  offsetPt: number,
): RowTrack {
  const rowPlacement: FlowBlockPlacement = {
    ...context.placement,
    cursor: { ...context.placement.cursor, yPt: context.placement.cursor.yPt + offsetPt },
  };
  return (_candidate, remaining) => {
    const laidOut = layoutTable({
      ...source.input,
      id: `${source.input.id}:float-probe:${context.page.occurrenceId}:${remaining.logicalRowIndex}`,
      rows: [remaining],
    }, rowPlacement, context.services).layout;
    const row = laidOut.rows[0];
    if (!row) throw new LayoutInvariantError('INVALID_REFERENCE', 'row track lost its row');
    return { input: remaining, row };
  };
}

/** The geometry a row is materialized in, independent of its cells'
 * content: its box and each cell's box (a merge owner's reaching the merge's
 * last row). Equal keys place one candidate's content identically, every
 * vAlign offset and block position included (table.ts materializeTableRow
 * reads nothing else of the tracks for them). */
function rowTrackKey(row: TableRowLayout): string {
  return JSON.stringify([row.flowBounds, row.cells.map((cell) => cell.flowBounds)]);
}

/**
 * The row's own §17.4.57 positioned children given final frames, their
 * anchor paragraphs re-acquired around them, to an exact fixed point. Every
 * anchor and column frame is read from the candidate as `track` lays it out
 * (the caller's final tracks of the row, or an estimate the caller checks
 * against them). `seed`, a previous resolution of these children in that
 * caller's solve, only starts the iteration: the accepted row is still the
 * one whose resolution, read from its own track, is the one it was
 * re-acquired with. Starting from the unwrapped row instead, a centered or
 * bottom-aligned cell whose wrapped content fills its track would approach
 * that fixed point without reaching it (each wrap moves the anchor a part of
 * the way).
 */
function finalFrameRow(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  track: RowTrack,
  context: TableFragmentContext,
  registry: readonly FloatRegistryEntryPt[],
  nextParagraphId: number,
  cursor: TableFragmentCursor,
  ownsAnchorStart: (
    occurrence: RetainedTableAcquisition['floatingTables'][number],
  ) => boolean,
  seed: readonly ResolvedFloatingTablePlacementLayout[] = [],
): Readonly<{
  row: TableRowLayoutInput;
  resolved: readonly ResolvedFloatingTablePlacementLayout[];
  registry: readonly FloatRegistryEntryPt[];
  nextParagraphId: number;
}> {
  const frames = context.floatingTableFrames;
  const reacquire = context.reacquirePageDependentBlock;
  const sourceRow = sourceRowFor(source, row.logicalRowIndex);
  if (!frames || !reacquire || !sourceRow) {
    return { row, resolved: [], registry, nextParagraphId };
  }
  // The row's own children, in source order (cell ids are unique).
  const { byHostCell } = floatingTableIndexOf(source);
  const occurrences = sourceRow.cells
    .flatMap((cell) => byHostCell.get(cell.id) ?? [])
    .sort((left, right) => left - right)
    .map((index) => source.floatingTables[index]!)
    .filter((occurrence) => ownsAnchorStart(occurrence));
  if (occurrences.length === 0) return { row, resolved: [], registry, nextParagraphId };
  const requiredAnchorByCell = new Map<string, number>();
  for (const occurrence of occurrences) {
    requiredAnchorByCell.set(
      occurrence.hostCellId,
      Math.min(
        requiredAnchorByCell.get(occurrence.hostCellId) ?? Number.POSITIVE_INFINITY,
        occurrence.anchorBlockIndex,
      ),
    );
  }

  const translation = context.finalPlacementTranslationPt ?? { xPt: 0, yPt: 0 };
  const placementFor = (
    occurrence: RetainedTableAcquisition['floatingTables'][number],
    laidOut: TableRowLayout,
    rowInput: TableRowLayoutInput,
  ): FloatingTablePlacementLayout | null => {
    // A track may lay out a fragment's cut of the row, whose cells carry
    // fragment ids: cells correspond by position, as the source row's do.
    const cellIndex = sourceRow.cells.findIndex((cell) => cell.id === occurrence.hostCellId);
    const laidOutCell = laidOut.cells[cellIndex];
    const selectedCell = rowInput.cells[cellIndex];
    const blockIndex = selectedCell?.blocks.findIndex((block) => (
      block.sourceBlockIndex === occurrence.anchorBlockIndex
    )) ?? -1;
    const anchorBlock = blockIndex < 0 ? undefined : laidOutCell?.blocks[blockIndex];
    const child = source.nestedById[occurrence.tableId]?.layout;
    if (!laidOutCell || !anchorBlock || !child) return null;
    return Object.freeze({
      kind: 'floating-table-placement' as const,
      occurrenceId: [
        context.page.occurrenceId,
        occurrence.hostCellId,
        occurrence.sourceBlockIndex,
        occurrence.tableId,
      ].join(':'),
      ownership,
      physicalPageIndex: context.page.physicalPageIndex,
      displayPageNumber: context.page.displayPageNumber,
      ...occurrence,
      columnBounds: Object.freeze({
        xPt: laidOutCell.contentBounds.xPt + translation.xPt,
        yPt: laidOutCell.contentBounds.yPt + translation.yPt,
        widthPt: laidOutCell.contentBounds.widthPt,
        heightPt: laidOutCell.contentBounds.heightPt,
      }),
      anchorBounds: Object.freeze({
        xPt: laidOutCell.contentBounds.xPt + translation.xPt,
        yPt: laidOutCell.flowBounds.yPt + anchorBlock.offsetPt + translation.yPt,
        widthPt: anchorBlock.layout.flowBounds.widthPt,
        heightPt: anchorBlock.layout.flowBounds.heightPt,
      }),
      child,
    });
  };

  const resolveCandidate = (candidate: TableRowLayoutInput) => {
    const laidOut = track(
      candidate,
      remainingRowAtCursor(source, candidate, cursor, context, requiredAnchorByCell),
    );
    let transaction = beginFloatingTablePlacementTransaction(
      registry,
      nextParagraphId,
      context.floatingTableRegistry?.coordinateSpace ?? 'logical-page-points',
      context.floatingTableRegistry?.flowDomainId ?? source.input.flowDomainId,
    );
    const resolved: ResolvedFloatingTablePlacementLayout[] = [];
    for (const occurrence of occurrences) {
      // Every owned occurrence gets its final frame, a text/text one too: no
      // later step places an unresolved nested occurrence on the page (only
      // resolved placements are painted). A text axis keeps the offset its
      // acquisition registered against the cell flow (acquiredTextOffsetPt),
      // so the anchor paragraphs' acquired wrap is unchanged.
      const placement = placementFor(occurrence, laidOut.row, laidOut.input);
      if (!placement) continue;
      const resolveWith = (child: TableLayout) => resolveFloatingTablePlacementInTransaction(
        child === placement.child ? placement : Object.freeze({ ...placement, child }),
        {
          page: frames.page,
          margin: frames.margin,
          text: {
            xPt: placement.columnBounds?.xPt ?? placement.anchorBounds.xPt,
            yPt: placement.anchorBounds.yPt,
            widthPt: placement.columnBounds?.widthPt ?? placement.anchorBounds.widthPt,
            heightPt: placement.anchorBounds.heightPt,
          },
        },
        transaction,
      );
      const resolution = resolveFinalFrameChild(source, occurrence, context, resolveWith);
      resolved.push(resolution.placement);
      transaction = resolution.transaction;
    }
    return { resolved: Object.freeze(resolved), transaction };
  };
  const reacquireCandidate = (
    resolved: readonly ResolvedFloatingTablePlacementLayout[],
  ): TableRowLayoutInput => ({
    ...row,
    cells: row.cells.map((cell, logicalCellIndex) => ({
      ...cell,
      blocks: cell.blocks.map((block) => {
        const exclusions = resolved.filter((placement) => (
          placement.source.hostCellId === cell.id
          && placement.source.anchorBlockIndex === block.sourceBlockIndex
        )).map((placement) => Object.freeze({
          xPt: placement.exclusionBounds.xPt - placement.source.anchorBounds.xPt,
          yPt: placement.exclusionBounds.yPt - placement.source.anchorBounds.yPt,
          widthPt: placement.exclusionBounds.widthPt,
          heightPt: placement.exclusionBounds.heightPt,
        }));
        if (exclusions.length === 0 || block.layout.kind !== 'paragraph') return block;
        return {
          ...block,
          layout: reacquire({
            logicalRowIndex: row.logicalRowIndex,
            logicalCellIndex,
            sourceBlockIndex: block.sourceBlockIndex,
            ownership,
            page: context.page,
            acquired: block.layout,
            floatingTableExclusions: Object.freeze(exclusions),
          }),
        };
      }),
    })),
  });
  const convergenceKey = (
    candidate: TableRowLayoutInput,
    resolved: readonly ResolvedFloatingTablePlacementLayout[],
  ) => JSON.stringify({
    blocks: candidate.cells.map((cell) => cell.blocks.map((block) => ({
      sourceBlockIndex: block.sourceBlockIndex,
      layout: block.layout,
    }))),
    placements: resolved,
  });

  // A seed resolution of children this row no longer owns is not used.
  const owned = new Set(occurrences.map(occurrenceSelectionKey));
  const seeded = seed.filter((placement) => owned.has(occurrenceSelectionKey(placement.source)));
  let initialResolved: readonly ResolvedFloatingTablePlacementLayout[] = seeded;
  if (seeded.length === 0) {
    initialResolved = resolveCandidate(row).resolved;
    if (initialResolved.length === 0) return { row, resolved: [], registry, nextParagraphId };
  }
  type Pass = Readonly<{
    candidate: TableRowLayoutInput;
    resolution: ReturnType<typeof resolveCandidate>;
    state: string;
  }>;
  try {
    const result = convergeExactState<Pass>({
      seedState: convergenceKey(row, initialResolved),
      step: (previous) => {
        const candidate = reacquireCandidate(
          previous?.resolution.resolved ?? initialResolved,
        );
        const resolution = resolveCandidate(candidate);
        return Object.freeze({
          candidate,
          resolution,
          state: convergenceKey(candidate, resolution.resolved),
        });
      },
      stateOf: (pass) => pass.state,
      // Resource guard only; exact equality/cycle detection determines
      // correctness and no last candidate is accepted on exhaustion.
      limit: 16,
    }).value;
    return {
      row: result.candidate,
      resolved: result.resolution.resolved,
      registry: Object.freeze([
        ...result.resolution.transaction.base,
        ...result.resolution.transaction.delta,
      ]),
      nextParagraphId: result.resolution.transaction.nextParagraphId,
    };
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `floating table final-frame reflow did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
}

function selectedOwnsOccurrence(
  source: RetainedTableAcquisition,
  selection: SelectedRow,
  occurrence: Pick<FloatingTablePlacementLayout, 'hostCellId' | 'anchorBlockIndex'>,
): boolean {
  const sourceRow = sourceRowFor(source, selection.logicalRowIndex);
  const cellIndex = sourceRow?.cells.findIndex(
    (cell) => cell.id === occurrence.hostCellId,
  ) ?? -1;
  return cellIndex >= 0 && (selection.ranges[cellIndex]?.some((range) => (
    range.blockIndex === occurrence.anchorBlockIndex
      && (range.kind === 'whole'
        || (range.kind === 'paragraph' && range.lineStart === 0)
        || (range.kind === 'nested-table' && range.childFragmentIndex === 0))
  )) ?? false);
}

function occurrenceSelectionKey(
  occurrence: Pick<FloatingTablePlacementLayout, 'hostCellId' | 'sourceBlockIndex' | 'tableId'>,
): string {
  return `${occurrence.hostCellId}:${occurrence.sourceBlockIndex}:${occurrence.tableId}`;
}

function selectedOccurrenceKeys(
  source: RetainedTableAcquisition,
  selection: SelectedRow,
): ReadonlySet<string> {
  return new Set(source.floatingTables.filter((occurrence) => (
    selectedOwnsOccurrence(source, selection, occurrence)
  )).map(occurrenceSelectionKey));
}

function sameStringSet(left: ReadonlySet<string>, right: ReadonlySet<string>): boolean {
  return left.size === right.size && [...left].every((item) => right.has(item));
}

function selectedWholeRow(
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  fragmentIndex = 0,
  clipAtPageEnd = false,
  resolvedFloatingTables: readonly ResolvedFloatingTablePlacementLayout[] = [],
): SelectedRow {
  return {
    input: row,
    logicalRowIndex: row.logicalRowIndex,
    fragmentIndex,
    ownership,
    ranges: rowRanges(row),
    ...(clipAtPageEnd ? { clipAtPageEnd: true } : {}),
    ...(resolvedFloatingTables.length ? { resolvedFloatingTables } : {}),
  };
}

function paragraphSlice(
  paragraph: ParagraphLayout,
  start: number,
  end: number,
): ParagraphLayout {
  return sliceParagraphLayout(paragraph, {
    lineStart: start,
    lineEnd: end,
    continuesFromPrevious: start > 0,
    continuesOnNext: end < paragraph.lines.length,
  });
}

function selectParagraph(
  paragraph: ParagraphLayout,
  sourceBlock: TableCellBlockInput,
  start: number,
  selectedBlocks: readonly TableCellBlockInput[],
  availableHeightPt: number,
  freshAvailableHeightPt: number,
  canGainPageSpace: boolean,
): Readonly<{
  block: TableCellBlockInput | null;
  range: BlockContinuationRange | null;
  lineEnd: number;
  advancePt: number;
}> {
  let selected: ParagraphLayout | null = null;
  let lineEnd = start;
  for (let candidateEnd = start + 1; candidateEnd <= paragraph.lines.length; candidateEnd += 1) {
    const candidate = paragraphSlice(paragraph, start, candidateEnd);
    const candidateBlock = { ...sourceBlock, layout: candidate };
    if (measureTableCellBlockFlowHeightPt([...selectedBlocks, candidateBlock])
      > availableHeightPt + EPSILON_PT) break;
    lineEnd = candidateEnd;
  }
  if (lineEnd === start) return { block: null, range: null, lineEnd: start, advancePt: 0 };
  const canRelocate = start === 0 && (selectedBlocks.length > 0 || canGainPageSpace);
  // §17.3.1.14 / §17.3.1.44 apply to cell paragraphs too. The selected
  // line count is the retained paragraph's count; no substitute-font or
  // manual-break-specific compatibility rule changes it here.
  if (sourceBlock.keepLines === true
    && lineEnd < paragraph.lines.length
    && canRelocate
    && measureTableCellBlockFlowHeightPt([{ ...sourceBlock, layout: paragraph }])
      <= freshAvailableHeightPt + EPSILON_PT) {
    return { block: null, range: null, lineEnd: start, advancePt: 0 };
  }
  for (;;) {
    const widow = adjustForWidowOrphan({
      widowControl: sourceBlock.widowControl === true,
      start,
      end: lineEnd,
      totalLines: paragraph.lines.length,
      canRelocate,
    });
    if (widow.kind === 'relocate') {
      return { block: null, range: null, lineEnd: start, advancePt: 0 };
    }
    if (widow.kind !== 'dropLastLine') break;
    lineEnd -= 1;
  }
  selected = paragraphSlice(paragraph, start, lineEnd);
  return {
    block: { ...sourceBlock, layout: selected },
    range: { kind: 'paragraph', blockIndex: sourceBlock.sourceBlockIndex, lineStart: start, lineEnd },
    lineEnd,
    advancePt: selected.advancePt,
  };
}

function selectCell(
  source: RetainedTableAcquisition,
  cell: TableCellLayoutInput,
  cursor: TableCellFragmentCursor,
  availableContentHeightPt: number,
  freshAvailableContentHeightPt: number,
  canGainPageSpace: boolean,
  context: TableFragmentContext,
): SelectedCell {
  if (cell.verticalMerge === 'continue') {
    return { input: cell, range: [], next: cursor, complete: true };
  }
  if (cell.verticalText) {
    // ECMA-376 §17.4.72: rotated lines run along the row, so the block axis
    // is the cell width and cannot be sliced across a page break. The first
    // fragment owns the whole rotated content.
    const atStart = cursor.blockIndex === 0 && cursor.paragraphLineStart === 0;
    return {
      input: { ...cell, blocks: atStart ? cell.blocks : [] },
      range: atStart
        ? cell.blocks.map((block) => ({ kind: 'whole' as const, blockIndex: block.sourceBlockIndex }))
        : [],
      next: Object.freeze({
        blockIndex: cell.blocks.length,
        paragraphLineStart: 0,
        nestedCursor: null,
        nestedFragmentIndex: 0,
      }),
      complete: true,
    };
  }
  const blocks: TableCellBlockInput[] = [];
  const range: BlockContinuationRange[] = [];
  let blockIndex = cursor.blockIndex;
  let paragraphLineStart = cursor.paragraphLineStart;
  let nestedCursor = cursor.nestedCursor;
  let nestedFragmentIndex = cursor.nestedFragmentIndex;

  while (blockIndex < cell.blocks.length) {
    const sourceBlock = cell.blocks[blockIndex]!;
    const child = sourceBlock.layout;
    if (child.kind === 'paragraph') {
      if (sourceBlock.structuralTrailing) {
        blocks.push(sourceBlock);
        range.push({ kind: 'whole', blockIndex: sourceBlock.sourceBlockIndex });
        blockIndex += 1;
        paragraphLineStart = 0;
        continue;
      }
      // A retained paragraph may legitimately own no paintable lines. It is
      // still a source block and must advance the cell cursor; asking the line
      // slicer for a first line can never make progress from an empty array.
      if (child.lines.length === 0) {
        if (measureTableCellBlockFlowHeightPt([...blocks, sourceBlock])
          > availableContentHeightPt + EPSILON_PT) break;
        blocks.push(sourceBlock);
        range.push({ kind: 'whole', blockIndex: sourceBlock.sourceBlockIndex });
        blockIndex += 1;
        paragraphLineStart = 0;
        continue;
      }
      const selected = selectParagraph(
        child,
        sourceBlock,
        paragraphLineStart,
        blocks,
        availableContentHeightPt,
        freshAvailableContentHeightPt,
        canGainPageSpace,
      );
      if (!selected.block || !selected.range) break;
      blocks.push({ ...selected.block, ...(sourceBlock.structuralTrailing
        ? { structuralTrailing: true }
        : {}) });
      range.push(selected.range);
      if (selected.lineEnd < child.lines.length) {
        paragraphLineStart = selected.lineEnd;
        break;
      }
      blockIndex += 1;
      paragraphLineStart = 0;
      continue;
    }

    const nested = source.nestedById[child.id];
    if (nested) {
      const remainingPt = Math.max(
        0,
        availableContentHeightPt - measureTableCellBlockFlowHeightPt(blocks),
      );
      const nestedResult = takeTableFragment(
        nested,
        nestedCursor ?? startTableFragmentCursor(),
        nestedTableFragmentContext(source, cell, context, remainingPt, sourceBlock),
      );
      if (!nestedResult.fragment) break;
      blocks.push({ layout: nestedResult.fragment, sourceBlockIndex: sourceBlock.sourceBlockIndex });
      range.push({
        kind: 'nested-table',
        blockIndex: sourceBlock.sourceBlockIndex,
        childFragmentIndex: nestedFragmentIndex,
      });
      if (nestedResult.nextCursor) {
        nestedCursor = nestedResult.nextCursor;
        nestedFragmentIndex += 1;
        break;
      }
      blockIndex += 1;
      nestedCursor = null;
      nestedFragmentIndex = 0;
      continue;
    }

    if (measureTableCellBlockFlowHeightPt([...blocks, sourceBlock])
      > availableContentHeightPt + EPSILON_PT) break;
    blocks.push(sourceBlock);
    range.push({ kind: 'whole', blockIndex: sourceBlock.sourceBlockIndex });
    blockIndex += 1;
  }

  const complete = blockIndex >= cell.blocks.length;
  return {
    input: { ...cell, blocks },
    range,
    next: Object.freeze({ blockIndex, paragraphLineStart, nestedCursor, nestedFragmentIndex }),
    complete,
  };
}

function partialRow(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  cursor: TableFragmentCursor,
  availableHeightPt: number,
  freshAvailableHeightPt: number,
  canGainPageSpace: boolean,
  context: TableFragmentContext,
): Readonly<{
  selected: SelectedRow | null;
  next: TableFragmentCursor;
  complete: boolean;
}> {
  const cellCursors = row.cells.map((_, index) => cursor.cells[index] ?? emptyCellCursor());
  // Library cut model: each fragment is materialized as its own grid in which
  // every cell keeps both §17.4.68 margins, a projected segment-opening owner
  // (table.ts projectedMergeRole) included, so each fragment reserves them.
  // Splitting one cell's margins across fragments is not modeled.
  const verticalInsetsPt = Math.max(0, ...row.cells.map((cell) => (
    cell.margins.topPt + cell.margins.bottomPt
  )));
  // A one-row fragment owns both outer cell-spacing bands. Reserve them before
  // selecting legal child boundaries so layoutTable cannot grow past the page.
  const spacingInsetsPt = Math.max(0, row.cellSpacingPt) * 2;
  // A continued row is materialized as an auto-height, one-row table fragment.
  // Reserve the same page-local collapsed top/bottom half-rules that layoutTable
  // will add to that track after the legal child boundary has been selected.
  const fragmentRow: TableRowLayoutInput = {
    ...row,
    heightPt: null,
    heightRule: 'auto',
  };
  const boundaryInsetsPt = tableRowBoundaryFootprintsPt({
    ...source.input,
    rows: [fragmentRow],
  })[0] ?? 0;
  const availableContentHeightPt = Math.max(
    0,
    availableHeightPt - verticalInsetsPt - spacingInsetsPt - boundaryInsetsPt,
  );
  const freshAvailableContentHeightPt = Math.max(
    0,
    freshAvailableHeightPt - verticalInsetsPt - spacingInsetsPt - boundaryInsetsPt,
  );
  const selectedCells = row.cells.map((cell, index) => selectCell(
    source,
    cell,
    cellCursors[index]!,
    availableContentHeightPt,
    freshAvailableContentHeightPt,
    canGainPageSpace,
    context,
  ));
  const cellMadeProgress = (cell: SelectedCell, index: number) => (
    cell.next.blockIndex !== cellCursors[index]?.blockIndex
    || cell.next.paragraphLineStart !== cellCursors[index]?.paragraphLineStart
    || cell.next.nestedFragmentIndex !== cellCursors[index]?.nestedFragmentIndex
  );
  // The compatibility authority owns the cross-cell cut selection that
  // ECMA-376 §17.4.6 leaves undefined. This call site deliberately checks only
  // paragraph children; nested tables retain their own fragmentation path.
  const hasUnfinishedParagraphWithoutProgress = selectedCells.some((cell, index) => (
    !cell.complete
    && !cellMadeProgress(cell, index)
    && row.cells[index]?.blocks[cellCursors[index]?.blockIndex ?? 0]?.layout.kind === 'paragraph'
  ));
  if (wordRelocatesParallelParagraphRowCut({
    compatibility: context.compatibility,
    hasUnfinishedParagraphWithoutProgress,
  })) {
    return { selected: null, next: cursor, complete: false };
  }
  const madeProgress = selectedCells.some(cellMadeProgress);
  if (!madeProgress) return { selected: null, next: cursor, complete: false };

  const complete = selectedCells.every((cell) => cell.complete);
  if (complete && cursor.rowFragmentIndex === 0) {
    return {
      selected: selectedWholeRow(row, 'source'),
      next: Object.freeze({
        rowIndex: cursor.rowIndex + 1,
        rowFragmentIndex: 0,
        cells: Object.freeze([]),
      }),
      complete: true,
    };
  }
  // Reaching this branch means content genuinely continues from or onto another
  // fragment: a fully retained first fragment returned as a whole row above.
  // Authored exact/atLeast height constrains that logical row once, not every
  // continuation, so fragment-local tracks must derive from retained content.
  const fragmentInput: TableRowLayoutInput = {
    ...fragmentRow,
    id: `${row.id}:fragment:${cursor.rowFragmentIndex}`,
    heightPt: null,
    heightRule: 'auto',
    cells: selectedCells.map((cell, index) => ({
      ...cell.input,
      id: `${cell.input.id}:fragment:${cursor.rowFragmentIndex}:${index}`,
    })),
  };
  return {
    selected: {
      input: fragmentInput,
      logicalRowIndex: row.logicalRowIndex,
      fragmentIndex: cursor.rowFragmentIndex,
      ownership: 'source',
      ranges: selectedCells.map((cell) => cell.range),
    },
    next: Object.freeze({
      rowIndex: complete ? cursor.rowIndex + 1 : cursor.rowIndex,
      rowFragmentIndex: complete ? 0 : cursor.rowFragmentIndex + 1,
      cells: complete ? Object.freeze([]) : Object.freeze(selectedCells.map((cell) => cell.next)),
    }),
    complete,
  };
}

function materializeFragment(
  source: RetainedTableAcquisition,
  selected: readonly SelectedRow[],
  context: TableFragmentContext,
): TableFragmentLayout {
  const fragmentInput: TableLayoutInput = {
    ...source.input,
    id: `${source.input.id}:fragment:${context.page.occurrenceId}`,
    rows: selected.map((row) => row.input),
  };
  const laidOut = layoutTable(fragmentInput, context.placement, context.services).layout;
  const rows = laidOut.rows.map((row, rowIndex): TableRowFragmentLayout => {
    const selection = selected[rowIndex]!;
    return Object.freeze({
      ...row,
      logicalRowIndex: selection.logicalRowIndex,
      fragmentIndex: selection.fragmentIndex,
      ownership: selection.ownership,
      occurrenceId: context.page.occurrenceId,
      physicalPageIndex: context.page.physicalPageIndex,
      displayPageNumber: context.page.displayPageNumber,
      cells: Object.freeze(row.cells.map((cell, cellIndex): TableCellFragmentLayout => {
        const verticalMerge = selection.input.cells[cellIndex]?.verticalMerge ?? 'none';
        const sourceCell = selection.input.cells[cellIndex];
        const ownsRestartInFragment = verticalMerge === 'continue' && selected
          .slice(0, rowIndex)
          .some((earlier) => earlier.input.cells.some((candidate) => (
            candidate.verticalMerge === 'restart'
            && candidate.columnStart === sourceCell?.columnStart
            && candidate.columnSpan === sourceCell?.columnSpan
          )));
        return Object.freeze({
          ...cell,
          contentRanges: Object.freeze([...(selection.ranges[cellIndex] ?? [])]),
          ...(verticalMerge === 'continue' && !ownsRestartInFragment
            ? { visualMergeOwnership: 'continuation' as const }
            : {}),
        });
      })),
    });
  });
  const floatingTables = selected.flatMap((selection, rowIndex) => {
    const sourceRow = sourceRowFor(source, selection.logicalRowIndex);
    if (!sourceRow) return [];
    return source.floatingTables.flatMap((occurrence): FloatingTablePlacementLayout[] => {
      const logicalCellIndex = sourceRow.cells.findIndex((cell) => cell.id === occurrence.hostCellId);
      if (logicalCellIndex < 0) return [];
      const ownsAnchorStart = selection.ranges[logicalCellIndex]?.some((range) => (
        range.blockIndex === occurrence.anchorBlockIndex
          && (range.kind === 'whole'
            || (range.kind === 'paragraph' && range.lineStart === 0))
      )) ?? false;
      if (!ownsAnchorStart) return [];

      const selectedCell = selection.input.cells[logicalCellIndex];
      const laidOutCell = rows[rowIndex]?.cells[logicalCellIndex];
      const anchorBlockOffset = selectedCell?.blocks.findIndex((block) => (
        block.sourceBlockIndex === occurrence.anchorBlockIndex
      )) ?? -1;
      const anchorBlock = anchorBlockOffset < 0
        ? undefined : laidOutCell?.blocks[anchorBlockOffset];
      const child = source.nestedById[occurrence.tableId]?.layout;
      if (!laidOutCell || !anchorBlock || !child) {
        throw new Error('Floating table occurrence references missing retained layout data');
      }
      const anchorBounds = Object.freeze({
        xPt: laidOutCell.contentBounds.xPt,
        yPt: laidOutCell.flowBounds.yPt + anchorBlock.offsetPt,
        widthPt: anchorBlock.layout.flowBounds.widthPt,
        heightPt: anchorBlock.layout.flowBounds.heightPt,
      });
      return [Object.freeze({
        kind: 'floating-table-placement' as const,
        occurrenceId: [
          context.page.occurrenceId,
          occurrence.hostCellId,
          occurrence.sourceBlockIndex,
          occurrence.tableId,
        ].join(':'),
        ownership: selection.ownership,
        physicalPageIndex: context.page.physicalPageIndex,
        displayPageNumber: context.page.displayPageNumber,
        ...occurrence,
        ...(occurrence.positioning.widthBasis === 'host-cell-content'
          ? { columnBounds: Object.freeze({ ...laidOutCell.contentBounds }) } : {}),
        anchorBounds,
        child,
      })];
    });
  });
  const resolvedFloatingTables = Object.freeze(selected.flatMap(
    (selection) => selection.resolvedFloatingTables ?? [],
  ));
  const resolvedOccurrenceIds = new Set(
    resolvedFloatingTables.map((placement) => placement.occurrenceId),
  );
  // Column measurement is acquisition-owned. A fragment may rebuild row and
  // border geometry, but must retain the one authoritative width vector.
  const clipAtPageEnd = selected.some((row) => row.clipAtPageEnd === true);
  const clippedHeightPt = clipAtPageEnd
    ? Math.min(laidOut.advancePt, context.availableHeightPt)
    : laidOut.advancePt;
  const flowBounds = clipAtPageEnd
    ? { ...laidOut.flowBounds, heightPt: clippedHeightPt }
    : laidOut.flowBounds;
  return Object.freeze({
    ...laidOut,
    flowBounds,
    ...(clipAtPageEnd ? {
      unpaintedOverflowPt: Math.max(0, laidOut.advancePt - clippedHeightPt),
      inkBounds: flowBounds,
      clipBounds: flowBounds,
      advancePt: clippedHeightPt,
    } : {}),
    columnWidthsPt: source.layout.columnWidthsPt,
    rows: Object.freeze(rows),
    floatingTables: Object.freeze(floatingTables.filter(
      (placement) => !resolvedOccurrenceIds.has(placement.occurrenceId),
    )),
    resolvedFloatingTables,
    ...(context.floatingTableRegistry ? {
      resolvedFloatingTableCoordinateSpace: context.floatingTableRegistry.coordinateSpace,
    } : {}),
  });
}

function firstCellAnchorPastPageBand(
  fragment: TableFragmentLayout,
  pageBottomPt: number,
): number {
  for (let rowIndex = 0; rowIndex < fragment.rows.length; rowIndex += 1) {
    const row = fragment.rows[rowIndex]!;
    for (const cell of row.cells) {
      for (const block of cell.blocks) {
        const paragraph = block.layout;
        if (paragraph.kind !== 'paragraph') continue;
        for (const drawing of paragraph.drawings) {
          if (drawing.anchorLayer?.layoutInCell !== true
            || drawing.anchorLayer.cellContainment === true
            || drawing.anchorLayer.verticalOwnership !== 'host'
            || drawing.orientation === 'upright-physical') continue;
          const bottomPt = cell.contentBounds.yPt + block.offsetPt
            + drawing.flowBounds.yPt - paragraph.flowBounds.yPt
            + drawing.flowBounds.heightPt;
          if (bottomPt > pageBottomPt + EPSILON_PT) return rowIndex;
        }
      }
    }
  }
  return -1;
}

/** How a selection pass prepares the row at one fragment position
 * ({@link takeTableFragment}): `final`, the track a previous pass's
 * materialization gave that position (absent before one did), and `seed`,
 * the row placed there in that pass. `used` is told every track the row's
 * preparation reads, the last being the accepted one's. */
type FragmentRowFrame = Readonly<{
  final?: (candidate: TableRowLayoutInput) => TableRowLayout;
  seed?: TableRowLayoutInput;
  used: (row: TableRowLayout) => void;
}>;

/** What the latest materialization holding a fragment position left for the
 * next selection pass: its track there, and what that pass placed and
 * resolved there (seeds only). */
type FragmentPositionHint = Readonly<{
  final: (candidate: TableRowLayoutInput) => TableRowLayout;
  placed?: TableRowLayoutInput;
  cut?: TableRowLayoutInput;
  resolved: readonly ResolvedFloatingTablePlacementLayout[];
}>;

type FragmentPass = Readonly<{
  result: TableFragmentResult;
  /** The rows the result's fragment materializes, in position order. */
  selected: readonly SelectedRow[];
  /** Per position, the key of the track its preparation was accepted in. */
  used: ReadonlyMap<number, string>;
  placed: ReadonlyMap<number, TableRowLayoutInput>;
  cut: ReadonlyMap<number, TableRowLayoutInput>;
}>;

/** Selection passes of one fragment; a resource guard only. */
const FRAGMENT_TRACK_PASS_LIMIT = 16;

/**
 * The next fragment of a paginated table (library policy for the rows'
 * page-placed content, below).
 *
 * A row whose content is placed by its geometry — page-placed content below
 * its cells (rowPagePlacement) and its own §17.4.57 positioned children
 * (finalFrameRow) — is prepared before the fragment holding it is laid out,
 * so a selection pass reads its track from an estimate: the rows selected
 * above it (probeFragmentRow), or the row by itself below them. Neither is the
 * materialized fragment's track when a merge continues below the row (the
 * merge's deficit then lands lower, not on the row) or when the rows below
 * change its rule footprint, and a centered or bottom-aligned cell's content,
 * every anchor in it and so every wrap, sits where its track puts it.
 *
 * So the fragment is solved to an exact fixed point: a pass selects and
 * materializes the fragment as before; if every row prepared against a track
 * was accepted in the very track the materialization gives it (equal
 * {@link rowTrackKey}), the pass is the result. Otherwise the next pass is
 * run from scratch — page admission, repeated headers, cut reselection and
 * the ownership of continuations included — with each position prepared in
 * the latest materialization's track of it (table.ts laidOutTableTracks) and
 * seeded by what was placed there. Admission therefore uses the heights of
 * rows prepared in their final tracks, and the result's registry delta and
 * placements are its own pass's (an earlier pass leaves nothing). No
 * threshold is applied; exhaustion fails closed.
 *
 * Cost: a table without such rows takes one pass, unchanged. Otherwise a
 * pass is the selection (proportional to the selected rows, with the prefix
 * probes) plus one frame build proportional to the fragment; passes are
 * bounded by the guard, and every pass after the first is charged to the
 * session's acquisition budget (runtime-state.ts), as a speculative whole
 * layout nested in another is, since a nested table's fragments are solved
 * inside its parent's passes.
 */
export function takeTableFragment(
  source: RetainedTableAcquisition,
  cursor: TableFragmentCursor,
  context: TableFragmentContext,
): TableFragmentResult {
  let hints: ReadonlyMap<number, FragmentPositionHint> = new Map();
  for (let pass = 1; ; pass += 1) {
    const taken = takeTableFragmentPass(source, cursor, context, hints);
    const fragment = taken.result.fragment;
    if (!fragment || taken.used.size === 0) return taken.result;
    const keys = fragment.rows.map(rowTrackKey);
    // A position past the fragment (trimmed after selection) placed nothing.
    if ([...taken.used].every(([position, key]) => position >= keys.length || keys[position] === key)) {
      return taken.result;
    }
    if (pass >= FRAGMENT_TRACK_PASS_LIMIT) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `table fragment row tracks did not converge (${pass} passes)`,
      );
    }
    paragraphAcquisitionCacheOf(context.services)?.noteMiss();
    const tracks = laidOutTableTracks(
      { ...source.input, rows: taken.selected.map((selection) => selection.input) },
      fragment,
      context.placement,
    );
    // A position the latest fragment lacks keeps the hint of the last one
    // that held it.
    const next = new Map(hints);
    taken.selected.forEach((selection, position) => {
      next.set(position, Object.freeze({
        final: (candidate: TableRowLayoutInput) => tracks.row(position, candidate),
        placed: taken.placed.get(position),
        cut: taken.cut.get(position),
        resolved: selection.resolvedFloatingTables ?? [],
      }));
    });
    hints = next;
  }
}

function takeTableFragmentPass(
  source: RetainedTableAcquisition,
  cursor: TableFragmentCursor,
  context: TableFragmentContext,
  hints: ReadonlyMap<number, FragmentPositionHint>,
): FragmentPass {
  const selected: SelectedRow[] = [];
  const used = new Map<number, string>();
  const placedAt = new Map<number, TableRowLayoutInput>();
  const cutAt = new Map<number, TableRowLayoutInput>();
  const passOf = (result: TableFragmentResult): FragmentPass => Object.freeze({
    result, selected, used, placed: placedAt, cut: cutAt,
  });
  // The row at `position` is prepared in the latest materialization's track
  // of it, else in the estimate.
  const frameAt = (position: number, seed?: TableRowLayoutInput): FragmentRowFrame => ({
    final: hints.get(position)?.final,
    seed,
    used: (row) => { used.set(position, rowTrackKey(row)); },
  });
  const trackAt = (
    position: number,
    offsetPt: number,
    inputOf: (candidate: TableRowLayoutInput, remaining: TableRowLayoutInput) => TableRowLayoutInput,
  ): RowTrack => {
    const final = hints.get(position)?.final;
    const estimate = aloneRowTrack(source, context, offsetPt);
    return (candidate, remaining) => {
      let laidOut: ReturnType<RowTrack>;
      if (final) {
        const input = inputOf(candidate, remaining);
        laidOut = { input, row: final(input) };
      } else {
        laidOut = estimate(candidate, remaining);
      }
      used.set(position, rowTrackKey(laidOut.row));
      return laidOut;
    };
  };
  const wholeInput = (candidate: TableRowLayoutInput) => candidate;
  if (cursor.rowIndex >= source.input.rows.length) {
    return passOf({ fragment: null, nextCursor: null, requiresFreshPage: false });
  }

  const registrySnapshot = context.floatingTableRegistry;
  if (registrySnapshot
    && registrySnapshot.flowDomainId.length === 0) {
    throw new Error('Floating table registry coordinate/domain mismatch');
  }
  let floatRegistry = Object.freeze([
    ...(registrySnapshot?.entries ?? []),
  ]) as readonly FloatRegistryEntryPt[];
  let floatParagraphId = registrySnapshot?.nextParagraphId ?? 0;
  let availablePt = Math.max(0, context.availableHeightPt);
  const pageOriginTracks = startFragmentRowTracks(source);
  const headerCount = leadingHeaderCount(source.input);
  if (cursor.rowIndex >= headerCount && cursor.rowIndex > 0 && headerCount > 0) {
    for (let rowIndex = 0; rowIndex < headerCount; rowIndex += 1) {
      const position = selected.length;
      const hint = hints.get(position);
      const acquiredHeader = placeRowNestedContent(source, rowForOccurrence(
        source,
        source.input.rows[rowIndex]!,
        'repeated-header',
        context,
      ), () => selected.map((item) => item.input), 'repeated-header', context, pageOriginTracks,
      undefined, undefined, frameAt(position, hint?.placed));
      placedAt.set(position, acquiredHeader);
      const preparedHeader = finalFrameRow(
        source,
        acquiredHeader,
        'repeated-header',
        trackAt(position, context.availableHeightPt - availablePt, wholeInput),
        context,
        floatRegistry,
        floatParagraphId,
        startTableFragmentCursor(),
        () => true,
        hint?.resolved,
      );
      const header = preparedHeader.row;
      const heightPt = paginationRowHeightForOccurrence(source, header, rowIndex, context);
      if (heightPt > availablePt + EPSILON_PT) {
        return passOf({ fragment: null, nextCursor: cursor, requiresFreshPage: true });
      }
      selected.push(selectedWholeRow(
        header,
        'repeated-header',
        0,
        false,
        preparedHeader.resolved,
      ));
      floatRegistry = preparedHeader.registry;
      floatParagraphId = preparedHeader.nextParagraphId;
      availablePt -= heightPt;
    }
  }

  const repeatedHeaderHeightPt = context.availableHeightPt - availablePt;
  const freshSourceHeightPt = Math.max(
    0,
    context.freshPageHeightPt - repeatedHeaderHeightPt,
  );
  let nextCursor: TableFragmentCursor | null = cursor;
  let rowIndex = cursor.rowIndex;
  // Acquisition-time tracks of rows whose nested content is placed per page
  // (re-acquired anchor or text box paragraphs) are not their occurrence
  // tracks, so they never admit the remainder.
  const retainedRemainderFits = cursor.rowFragmentIndex === 0
    && cursor.cells.length === 0
    && !needsNestedPageOrigins(source)
    && source.layout.rows
      .slice(cursor.rowIndex)
      .reduce((heightPt, row) => heightPt + Math.max(0, row.heightPt), 0)
      <= availablePt + EPSILON_PT;
  let followsCompletedPartialRow = false;
  while (rowIndex < source.input.rows.length) {
    const ownership: TableFragmentOwnership = 'source';
    const rowCursor = rowIndex === cursor.rowIndex
      ? cursor
      : Object.freeze({ rowIndex, rowFragmentIndex: 0, cells: Object.freeze([]) });
    const occurrenceRow = rowForOccurrence(
      source,
      source.input.rows[rowIndex]!,
      ownership,
      context,
    );
    const position = selected.length;
    const hint = hints.get(position);
    const acquiredRow = placeRowNestedContent(
      source, occurrenceRow, () => selected.map((item) => item.input), ownership, context,
      pageOriginTracks, rowCursor, undefined, frameAt(position, hint?.placed),
    );
    placedAt.set(position, acquiredRow);
    const canTakeWhole = rowIndex !== cursor.rowIndex || cursor.rowFragmentIndex === 0;
    const preparedRow = canTakeWhole ? finalFrameRow(
      source,
      acquiredRow,
      ownership,
      trackAt(position, context.availableHeightPt - availablePt, wholeInput),
      context,
      floatRegistry,
      floatParagraphId,
      rowCursor,
      (occurrence) => {
        const cellIndex = acquiredRow.cells.findIndex(
          (cell) => cell.id === occurrence.hostCellId,
        );
        const anchorBlockOffset = acquiredRow.cells[cellIndex]?.blocks.findIndex(
          (block) => block.sourceBlockIndex === occurrence.anchorBlockIndex,
        ) ?? -1;
        if (anchorBlockOffset < 0) return false;
        const cellCursor = rowCursor.cells[cellIndex] ?? emptyCellCursor();
        return cellCursor.blockIndex < anchorBlockOffset
          || (cellCursor.blockIndex === anchorBlockOffset
            && cellCursor.paragraphLineStart === 0);
      },
      hint?.resolved,
    ) : {
      row: acquiredRow,
      resolved: Object.freeze([]),
      registry: floatRegistry,
      nextParagraphId: floatParagraphId,
    };
    const row = preparedRow.row;
    // If every retained physical track fits, admit the canonical table by those
    // tracks. A vMerge owner's contentHeightPt spans its following tracks and
    // must not be charged again. When the remainder does cross the boundary,
    // keep the conservative content-aware height so partial rows and page-local
    // merge continuation ownership are derived before materialization.
    const wholeHeightPt = retainedRemainderFits || followsCompletedPartialRow
      ? paginationRowTrackHeightForOccurrence(source, row, rowIndex, context)
      : paginationRowHeightForOccurrence(source, row, rowIndex, context);
    if (canTakeWhole) {
      if (wholeHeightPt <= availablePt + EPSILON_PT) {
        selected.push(selectedWholeRow(row, 'source', 0, false, preparedRow.resolved));
        floatRegistry = preparedRow.registry;
        floatParagraphId = preparedRow.nextParagraphId;
        availablePt -= wholeHeightPt;
        rowIndex += 1;
        nextCursor = rowIndex < source.input.rows.length
          ? Object.freeze({ rowIndex, rowFragmentIndex: 0, cells: Object.freeze([]) })
          : null;
        continue;
      }
    }

    if (row.cantSplit) {
      const selectedSourceRows = selected.some((item) => item.ownership === 'source');
      if (selectedSourceRows) break;
      const freshHeaderHeightPt = context.availableHeightPt - availablePt;
      const fitsFreshBand = wholeHeightPt + freshHeaderHeightPt
        <= context.freshPageHeightPt + EPSILON_PT;
      if (fitsFreshBand) {
        return passOf({ fragment: null, nextCursor: cursor, requiresFreshPage: true });
      }
      if (context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt) {
        return passOf({ fragment: null, nextCursor: cursor, requiresFreshPage: true });
      }
      // Compatibility-owned over-page cantSplit admission.
      if (wordClipsOverPageCantSplitRow({
        compatibility: context.compatibility,
        availableHeightPt: context.availableHeightPt,
        freshPageHeightPt: context.freshPageHeightPt,
        epsilonPt: EPSILON_PT,
      })) {
        selected.push(selectedWholeRow(row, 'source', 0, true, preparedRow.resolved));
        floatRegistry = preparedRow.registry;
        floatParagraphId = preparedRow.nextParagraphId;
        nextCursor = rowIndex + 1 < source.input.rows.length
          ? Object.freeze({ rowIndex: rowIndex + 1, rowFragmentIndex: 0, cells: Object.freeze([]) })
          : null;
        break;
      }
    // ECMA-376 §17.4.6 permits a row taller than a full page to continue;
      // only the explicit compatibility mode clips it.
    }

    if (canTakeWhole && wordRelocatesAuthoredHeightRowAtPageBoundary({
      compatibility: context.compatibility,
      heightRule: row.heightRule,
      repeatedHeader: row.repeatedHeader,
      authoredHeightPt: row.heightPt,
      availableHeightPt: availablePt,
      wholeHeightPt,
      freshAvailableHeightPt: freshSourceHeightPt,
      epsilonPt: EPSILON_PT,
    })) {
      if (selected.some((item) => item.ownership === 'source')) break;
      return passOf({ fragment: null, nextCursor: cursor, requiresFreshPage: true });
    }

    // Floating overflow is not defined by §17.4.57. The retained floating
    // adapter preserves the established row-boundary policy: after relocation
    // to a fresh band, one over-band row is emitted once instead of being
    // converted into synthetic line fragments. Ordinary tables keep the
    // specification-backed default split policy above.
    if (context.oversizedRowPolicy === 'atomic'
      && selected.every((item) => item.ownership === 'repeated-header')
      && context.availableHeightPt + EPSILON_PT >= context.freshPageHeightPt
      && wholeHeightPt > context.freshPageHeightPt + EPSILON_PT) {
      selected.push(selectedWholeRow(row, 'source', 0, false, preparedRow.resolved));
      floatRegistry = preparedRow.registry;
      floatParagraphId = preparedRow.nextParagraphId;
      nextCursor = rowIndex + 1 < source.input.rows.length
        ? Object.freeze({ rowIndex: rowIndex + 1, rowFragmentIndex: 0, cells: Object.freeze([]) })
        : null;
      break;
    }

    const canGainPageSpace = context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt
      || selected.some((item) => item.ownership === 'source');
    // The row is cut (or continued) here: its nested page-placed content is
    // placed against the occurrence this fragment selects, not the uncut row,
    // and so are its own positioned children's anchors.
    const cutOf = (candidate: TableRowLayoutInput): TableRowLayoutInput | null => {
      const probe = partialRow(
        source, candidate, rowCursor, availablePt, freshSourceHeightPt, canGainPageSpace, context,
      );
      return probe.selected && !(probe.complete && rowCursor.rowFragmentIndex === 0)
        ? probe.selected.input
        : null;
    };
    const continuing = rowCursor.rowFragmentIndex > 0;
    const cutRow = rowNeedsPageOrigins(source, occurrenceRow)
      ? placeRowNestedContent(
        source, occurrenceRow, () => selected.map((item) => item.input), ownership, context,
        pageOriginTracks, rowCursor, cutOf, frameAt(position, hint?.cut),
      )
      : acquiredRow;
    if (cutRow !== acquiredRow) cutAt.set(position, cutRow);
    const cutTrack = trackAt(
      position,
      context.availableHeightPt - availablePt,
      (candidate, remaining) => cutOf(candidate) ?? (continuing ? remaining : candidate),
    );
    let partial = partialRow(
      source, cutRow, rowCursor, availablePt,
      freshSourceHeightPt, canGainPageSpace, context,
    );
    let selectedPrepared: ReturnType<typeof finalFrameRow> | null = null;
    const visitedOwnershipStates = new Set<string>();
    while (partial.selected) {
      const transactionInputs = selectedOccurrenceKeys(source, partial.selected);
      const ownershipState = JSON.stringify([...transactionInputs].sort());
      if (visitedOwnershipStates.has(ownershipState)) {
        throw new Error('Floating table selected ownership did not converge');
      }
      visitedOwnershipStates.add(ownershipState);
      selectedPrepared = finalFrameRow(
        source,
        cutRow,
        ownership,
        cutTrack,
        context,
        floatRegistry,
        floatParagraphId,
        rowCursor,
        (occurrence) => transactionInputs.has(occurrenceSelectionKey(occurrence)),
        hint?.resolved,
      );
      const reselection = partialRow(
        source, selectedPrepared.row, rowCursor, availablePt,
        freshSourceHeightPt, canGainPageSpace, context,
      );
      if (!reselection.selected) {
        partial = reselection;
        break;
      }
      const reselectedInputs = selectedOccurrenceKeys(source, reselection.selected);
      partial = reselection;
      if (sameStringSet(transactionInputs, reselectedInputs)) break;
      selectedPrepared = null;
    }
    if (partial.selected && selectedPrepared === null) {
      throw new Error('Floating table selected ownership did not converge');
    }
    if (partial.selected) {
      const ownedResolved = selectedPrepared?.resolved ?? [];
      if (ownedResolved.some((placement) => (
        !selectedOwnsOccurrence(source, partial.selected!, placement.source)
      ))) {
        throw new Error('Floating table transaction included an unowned occurrence');
      }
      const baseRegistryLength = floatRegistry.length;
      const committedEntries = (selectedPrepared?.registry ?? floatRegistry)
        .slice(baseRegistryLength);
      selected.push({
        ...partial.selected,
        ...(ownedResolved.length
          ? { resolvedFloatingTables: Object.freeze(ownedResolved) }
          : {}),
      });
      floatRegistry = Object.freeze([...floatRegistry, ...committedEntries]);
      floatParagraphId += committedEntries.length;
      nextCursor = partial.next.rowIndex >= source.input.rows.length ? null : partial.next;
      if (partial.complete && partial.next.rowIndex < source.input.rows.length) {
        // A completed continuation can own vertically merged content whose
        // physical track is resolved only after following logical rows join the
        // same fragment. Keep admitting those rows by their retained tracks;
        // the canonical materialization below remains the final fit authority
        // and trims any over-admission at whole-row boundaries.
        availablePt = Math.max(0, availablePt - completedPartialRowTrackHeight(
          source,
          partial.selected.input,
          rowIndex,
          context,
        ));
        followsCompletedPartialRow = true;
        rowIndex = partial.next.rowIndex;
        continue;
      }
    }
    break;
  }

  const sourceRows = selected.filter((row) => row.ownership === 'source');
  if (sourceRows.length === 0) {
    const canProgressOnFreshPage = context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt;
    if (!canProgressOnFreshPage) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        'Table pagination cannot advance from a fresh page',
      );
    }
    return passOf({
      fragment: null,
      nextCursor: cursor,
      requiresFreshPage: true,
    });
  }
  let fragment = materializeFragment(source, selected, context);
  // §20.4.2.3 identifies the drawing as cell-owned; it does not specify this
  // page-cut choice. An overlapping drawing need not enlarge row flow height,
  // so a legal text cut can leave its image outside the body band. The bounded
  // Word choice and counterexamples are owned by WORD_CELL_OWNED_ANCHOR_PAGE_CUT.
  const pageBottomPt = context.placement.cursor.yPt + context.availableHeightPt;
  while (wordDefersCellOwnedAnchorPastPageBand({
    compatibility: context.compatibility,
    availableHeightPt: context.availableHeightPt,
    freshPageHeightPt: context.freshPageHeightPt,
    epsilonPt: EPSILON_PT,
  })) {
    const conflictIndex = firstCellAnchorPastPageBand(fragment, pageBottomPt);
    if (conflictIndex < 0) break;
    const conflict = selected[conflictIndex];
    const authoredRow = conflict && sourceRowFor(source, conflict.logicalRowIndex);
    if (conflict?.ownership !== 'source' || conflict.fragmentIndex !== 0
      || authoredRow?.heightRule !== 'atLeast' || authoredRow.cantSplit) break;
    if (selected.slice(0, conflictIndex).every((item) => item.ownership !== 'source')) {
      return passOf({ fragment: null, nextCursor: cursor, requiresFreshPage: true });
    }
    selected.splice(conflictIndex);
    nextCursor = Object.freeze({
      rowIndex: inputRowIndexOf(source, conflict.logicalRowIndex),
      rowFragmentIndex: 0,
      cells: Object.freeze([]),
    });
    fragment = materializeFragment(source, selected, context);
  }
  while (fragment.advancePt > context.availableHeightPt + EPSILON_PT) {
    const last = selected.at(-1);
    const sourceCount = selected.filter((row) => row.ownership === 'source').length;
    // Only a first-fragment source row is a legal trim boundary. Materializing
    // a fragment-truncated vMerge span relocates the owner's deficit into the
    // span's last row (table.ts resolveRowHeights), which can grow a trailing
    // partial row past the budget its lines were selected against. Such a row
    // has emitted nothing yet, so deferring it whole loses no content; a later
    // fragment (fragmentIndex > 0) or a non-source row would.
    const trimmableSourceRow = last?.ownership === 'source'
      && last.fragmentIndex === 0;
    if (!trimmableSourceRow || sourceCount <= 1) break;
    selected.pop();
    nextCursor = Object.freeze({
      rowIndex: inputRowIndexOf(source, last.logicalRowIndex),
      rowFragmentIndex: 0,
      cells: Object.freeze([]),
    });
    fragment = materializeFragment(source, selected, context);
  }
  if (fragment.advancePt > context.availableHeightPt + EPSILON_PT
    && context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt
    && fragment.advancePt <= context.freshPageHeightPt + EPSILON_PT) {
    return passOf({ fragment: null, nextCursor: cursor, requiresFreshPage: true });
  }
  return passOf({
    fragment,
    nextCursor,
    requiresFreshPage: false,
    floatingTablePlacements: fragment.resolvedFloatingTables,
    ...(registrySnapshot ? {
      floatingTableRegistryDelta: (() => {
        const selectedEntries = floatRegistry.slice(registrySnapshot.entries.length).filter((entry) => (
          fragment.resolvedFloatingTables.some(
            (placement) => placement.occurrenceId === entry.occurrenceId,
          )
        ));
        return floatingTableRegistryDelta(
          registrySnapshot,
          selectedEntries,
          registrySnapshot.nextParagraphId + selectedEntries.length,
        );
      })(),
    } : {}),
  });
}
