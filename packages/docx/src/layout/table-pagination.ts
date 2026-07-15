import type { RetainedTableAcquisition } from './table-acquisition.js';
import {
  beginFloatingTablePlacementTransaction,
  floatingTableRegistryDelta,
  resolveFloatingTablePlacementInTransaction,
} from './floating-table-transaction.js';
import { sliceParagraphLayout } from './paragraph.js';
import { layoutTable, measureTableCellBlockFlowHeightPt } from './table.js';
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
  /** A page-local paint role; the source w:vMerge value remains unchanged. */
  readonly visualMergeOwnership?: 'continuation';
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
  readonly rows: readonly TableRowFragmentLayout[];
  readonly floatingTables: readonly FloatingTablePlacementLayout[];
  readonly resolvedFloatingTables: readonly ResolvedFloatingTablePlacementLayout[];
  readonly floatingTableCoordinateSpace?: FloatRegistrySnapshotPt['coordinateSpace'];
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
  /** Reacquire only content whose destination page can change its geometry. */
  readonly reacquirePageDependentBlock?: (
    request: PageDependentTableBlockRequest,
  ) => ParagraphLayout | TableLayout;
}

export interface TableFragmentResult {
  readonly fragment: TableFragmentLayout | null;
  readonly nextCursor: TableFragmentCursor | null;
  readonly requiresFreshPage: boolean;
  readonly floatingTablePlacements?: readonly ResolvedFloatingTablePlacementLayout[];
  readonly floatingTableRegistryDelta?: FloatRegistryDeltaPt;
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

function leadingHeaderCount(input: TableLayoutInput): number {
  let count = 0;
  while (input.rows[count]?.repeatedHeader === true) count += 1;
  return count;
}

function paginationRowHeight(source: RetainedTableAcquisition, rowIndex: number): number {
  const row = source.layout.rows[rowIndex];
  if (!row) return 0;
  // ECMA-376 §17.4.80 requires exact content to remain inside the authored row
  // box. `layoutTable` has already resolved that track, including Word's largest
  // bottom-cell-padding addition ([MS-OI29500] 2.1.180(d)); content overflow is
  // paint-time clip ink and must not turn a fitting exact row into continuations.
  if (source.input.rows[rowIndex]?.heightRule === 'exact') return Math.max(0, row.heightPt);
  // A vertically merged owner's content requirement and its physical track can
  // land on different source rows. Considering both permits a legal boundary
  // through the span without rewriting restart/continue semantics.
  return Math.max(0, row.heightPt, row.contentHeightPt);
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

function ownsFinalFrameAxis(placement: FloatingTablePlacementLayout): boolean {
  const horizontal = placement.positioning.horzSpecified
    && (placement.positioning.horzAnchor === 'page'
      || placement.positioning.horzAnchor === 'margin');
  const vertical = placement.positioning.vertAnchor === 'page'
    || placement.positioning.vertAnchor === 'margin';
  return horizontal || vertical;
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
              const remaining = takeTableFragment(nested, cellCursor.nestedCursor, {
                ...context,
                availableHeightPt: context.freshPageHeightPt,
                placement: {
                  ...context.placement,
                  availableBounds: {
                    ...context.placement.availableBounds,
                    heightPt: context.freshPageHeightPt,
                  },
                },
              });
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

function finalFrameRow(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  rowOffsetPt: number,
  context: TableFragmentContext,
  registry: readonly FloatRegistryEntryPt[],
  nextParagraphId: number,
  cursor: TableFragmentCursor,
  ownsAnchorStart: (
    occurrence: RetainedTableAcquisition['floatingTables'][number],
  ) => boolean,
): Readonly<{
  row: TableRowLayoutInput;
  resolved: readonly ResolvedFloatingTablePlacementLayout[];
  registry: readonly FloatRegistryEntryPt[];
  nextParagraphId: number;
}> {
  const frames = context.floatingTableFrames;
  const reacquire = context.reacquirePageDependentBlock;
  const sourceRow = source.input.rows[row.logicalRowIndex];
  if (!frames || !reacquire || !sourceRow) {
    return { row, resolved: [], registry, nextParagraphId };
  }
  const occurrences = source.floatingTables.filter((occurrence) => (
    sourceRow.cells.some((cell) => cell.id === occurrence.hostCellId)
    && ownsAnchorStart(occurrence)
  ));
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

  const rowPlacement: FlowBlockPlacement = {
    ...context.placement,
    cursor: {
      ...context.placement.cursor,
      yPt: context.placement.cursor.yPt + rowOffsetPt,
    },
  };
  const remainingRow = remainingRowAtCursor(
    source, row, cursor, context, requiredAnchorByCell,
  );
  const provisional = layoutTable({
    ...source.input,
    id: `${source.input.id}:float-probe:${context.page.occurrenceId}:${row.logicalRowIndex}`,
    rows: [remainingRow],
  }, rowPlacement, context.services).layout;
  const translation = context.finalPlacementTranslationPt ?? { xPt: 0, yPt: 0 };
  const placementFor = (
    occurrence: RetainedTableAcquisition['floatingTables'][number],
    laidOut: TableLayout,
    rowInput: TableRowLayoutInput,
  ): FloatingTablePlacementLayout | null => {
    const cellIndex = rowInput.cells.findIndex((cell) => cell.id === occurrence.hostCellId);
    const laidOutCell = laidOut.rows[0]?.cells[cellIndex];
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
    const remainingCandidate = remainingRowAtCursor(
      source, candidate, cursor, context, requiredAnchorByCell,
    );
    const laidOut = candidate === row ? provisional : layoutTable({
      ...source.input,
      id: `${source.input.id}:float-converge:${context.page.occurrenceId}:${row.logicalRowIndex}`,
      rows: [remainingCandidate],
    }, rowPlacement, context.services).layout;
    let transaction = beginFloatingTablePlacementTransaction(
      registry,
      nextParagraphId,
      context.floatingTableRegistry?.coordinateSpace ?? 'logical-page-points',
      context.floatingTableRegistry?.flowDomainId ?? source.input.flowDomainId,
    );
    const resolved: ResolvedFloatingTablePlacementLayout[] = [];
    for (const occurrence of occurrences) {
      const placement = placementFor(occurrence, laidOut, remainingCandidate);
      if (!placement || (
        context.floatingTableRegistry?.coordinateSpace !== 'upright-physical-page-points'
        && !ownsFinalFrameAxis(placement)
      )) continue;
      const resolution = resolveFloatingTablePlacementInTransaction(placement, {
        page: frames.page,
        margin: frames.margin,
        text: {
          xPt: placement.columnBounds?.xPt ?? placement.anchorBounds.xPt,
          yPt: placement.anchorBounds.yPt,
          widthPt: placement.columnBounds?.widthPt ?? placement.anchorBounds.widthPt,
          heightPt: placement.anchorBounds.heightPt,
        },
      }, transaction);
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
    blocks: candidate.cells.flatMap((cell) => cell.blocks.map((block) => ({
      sourceBlockIndex: block.sourceBlockIndex,
      advancePt: block.layout.advancePt,
      lineCount: block.layout.kind === 'paragraph' ? block.layout.lines.length : null,
    }))),
    placements: resolved.map((placement) => ({
      occurrenceId: placement.occurrenceId,
      bounds: placement.bounds,
      exclusionBounds: placement.exclusionBounds,
      anchorBounds: placement.source.anchorBounds,
    })),
  });

  let candidate = row;
  let previousKey: string | null = null;
  for (let iteration = 0; iteration < 4; iteration += 1) {
    const resolution = resolveCandidate(candidate);
    if (resolution.resolved.length === 0) {
      return { row, resolved: [], registry, nextParagraphId };
    }
    const key = convergenceKey(candidate, resolution.resolved);
    if (previousKey === key) {
      return {
        row: candidate,
        resolved: resolution.resolved,
        registry: Object.freeze([
          ...resolution.transaction.base,
          ...resolution.transaction.delta,
        ]),
        nextParagraphId: resolution.transaction.nextParagraphId,
      };
    }
    previousKey = key;
    candidate = reacquireCandidate(resolution.resolved);
  }
  throw new Error('Floating table final-frame reflow did not converge');
}

function selectedOwnsOccurrence(
  source: RetainedTableAcquisition,
  selection: SelectedRow,
  occurrence: Pick<FloatingTablePlacementLayout, 'hostCellId' | 'anchorBlockIndex'>,
): boolean {
  const sourceRow = source.input.rows[selection.logicalRowIndex];
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
  sourceBlockIndex: number,
  start: number,
  selectedBlocks: readonly TableCellBlockInput[],
  availableHeightPt: number,
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
    const candidateBlock = { layout: candidate, sourceBlockIndex } as const;
    if (measureTableCellBlockFlowHeightPt([...selectedBlocks, candidateBlock])
      > availableHeightPt + EPSILON_PT) break;
    selected = candidate;
    lineEnd = candidateEnd;
  }
  if (!selected) return { block: null, range: null, lineEnd: start, advancePt: 0 };
  return {
    block: { layout: selected, sourceBlockIndex },
    range: { kind: 'paragraph', blockIndex: sourceBlockIndex, lineStart: start, lineEnd },
    lineEnd,
    advancePt: selected.advancePt,
  };
}

function selectCell(
  source: RetainedTableAcquisition,
  cell: TableCellLayoutInput,
  cursor: TableCellFragmentCursor,
  availableContentHeightPt: number,
  context: TableFragmentContext,
): SelectedCell {
  if (cell.verticalMerge === 'continue') {
    return { input: cell, range: [], next: cursor, complete: true };
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
      const selected = selectParagraph(
        child,
        sourceBlock.sourceBlockIndex,
        paragraphLineStart,
        blocks,
        availableContentHeightPt,
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
      const nestedResult = takeTableFragment(nested, nestedCursor ?? startTableFragmentCursor(), {
        ...context,
        availableHeightPt: remainingPt,
        placement: {
          ...context.placement,
          availableBounds: { ...context.placement.availableBounds, heightPt: remainingPt },
        },
      });
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
  context: TableFragmentContext,
): Readonly<{
  selected: SelectedRow | null;
  next: TableFragmentCursor;
  complete: boolean;
}> {
  const cellCursors = row.cells.map((_, index) => cursor.cells[index] ?? emptyCellCursor());
  const verticalInsetsPt = Math.max(0, ...row.cells.map((cell) => (
    cell.margins.topPt + cell.margins.bottomPt
  )));
  // A one-row fragment owns both outer cell-spacing bands. Reserve them before
  // selecting legal child boundaries so layoutTable cannot grow past the page.
  const spacingInsetsPt = Math.max(0, row.cellSpacingPt) * 2;
  const availableContentHeightPt = Math.max(
    0,
    availableHeightPt - verticalInsetsPt - spacingInsetsPt,
  );
  const selectedCells = row.cells.map((cell, index) => selectCell(
    source,
    cell,
    cellCursors[index]!,
    availableContentHeightPt,
    context,
  ));
  const madeProgress = selectedCells.some((cell, index) => (
    cell.next.blockIndex !== cellCursors[index]?.blockIndex
    || cell.next.paragraphLineStart !== cellCursors[index]?.paragraphLineStart
    || cell.next.nestedFragmentIndex !== cellCursors[index]?.nestedFragmentIndex
  ));
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
    ...row,
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
          floatingSourceBlocks: Object.freeze(source.floatingTables
            .filter((occurrence) => occurrence.hostCellId === cell.id)
            .map((occurrence) => Object.freeze({
              sourceBlockIndex: occurrence.sourceBlockIndex,
              tableId: occurrence.tableId,
            }))),
          ...(verticalMerge === 'continue' && !ownsRestartInFragment
            ? { visualMergeOwnership: 'continuation' as const }
            : {}),
        });
      })),
    });
  });
  const floatingTables = selected.flatMap((selection, rowIndex) => {
    const sourceRow = source.input.rows[selection.logicalRowIndex];
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
      floatingTableCoordinateSpace: context.floatingTableRegistry.coordinateSpace,
    } : {}),
  });
}

export function takeTableFragment(
  source: RetainedTableAcquisition,
  cursor: TableFragmentCursor,
  context: TableFragmentContext,
): TableFragmentResult {
  if (cursor.rowIndex >= source.input.rows.length) {
    return { fragment: null, nextCursor: null, requiresFreshPage: false };
  }

  const selected: SelectedRow[] = [];
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
  const headerCount = leadingHeaderCount(source.input);
  if (cursor.rowIndex >= headerCount && cursor.rowIndex > 0 && headerCount > 0) {
    for (let rowIndex = 0; rowIndex < headerCount; rowIndex += 1) {
      const acquiredHeader = rowForOccurrence(
        source,
        source.input.rows[rowIndex]!,
        'repeated-header',
        context,
      );
      const preparedHeader = finalFrameRow(
        source,
        acquiredHeader,
        'repeated-header',
        context.availableHeightPt - availablePt,
        context,
        floatRegistry,
        floatParagraphId,
        startTableFragmentCursor(),
        () => true,
      );
      const header = preparedHeader.row;
      const heightPt = paginationRowHeightForOccurrence(source, header, rowIndex, context);
      if (heightPt > availablePt + EPSILON_PT) {
        return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
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

  let nextCursor: TableFragmentCursor | null = cursor;
  let rowIndex = cursor.rowIndex;
  while (rowIndex < source.input.rows.length) {
    const ownership: TableFragmentOwnership = 'source';
    const acquiredRow = rowForOccurrence(
      source,
      source.input.rows[rowIndex]!,
      ownership,
      context,
    );
    const rowCursor = rowIndex === cursor.rowIndex
      ? cursor
      : Object.freeze({ rowIndex, rowFragmentIndex: 0, cells: Object.freeze([]) });
    const canTakeWhole = rowIndex !== cursor.rowIndex || cursor.rowFragmentIndex === 0;
    const preparedRow = canTakeWhole ? finalFrameRow(
      source,
      acquiredRow,
      ownership,
      context.availableHeightPt - availablePt,
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
    ) : {
      row: acquiredRow,
      resolved: Object.freeze([]),
      registry: floatRegistry,
      nextParagraphId: floatParagraphId,
    };
    const row = preparedRow.row;
    const wholeHeightPt = paginationRowHeightForOccurrence(source, row, rowIndex, context);
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
        return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
      }
      if (context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt) {
        return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
      }
      // [MS-OI29500] 2.1.120: Word starts an over-page cantSplit row on the
      // fresh page and clips its overflow instead of synthesizing a continuation.
      if (context.compatibility === 'word'
        && context.availableHeightPt + EPSILON_PT >= context.freshPageHeightPt) {
        selected.push(selectedWholeRow(row, 'source', 0, true, preparedRow.resolved));
        floatRegistry = preparedRow.registry;
        floatParagraphId = preparedRow.nextParagraphId;
        nextCursor = rowIndex + 1 < source.input.rows.length
          ? Object.freeze({ rowIndex: rowIndex + 1, rowFragmentIndex: 0, cells: Object.freeze([]) })
          : null;
        break;
      }
    // ECMA-376 §17.4.6 permits a row taller than a full page to continue;
      // only the documented Word compatibility mode clips it.
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

    let partial = partialRow(
      source, acquiredRow, rowCursor, availablePt, context,
    );
    let selectedPrepared: ReturnType<typeof finalFrameRow> | null = null;
    for (let ownershipPass = 0; partial.selected && ownershipPass < 4; ownershipPass += 1) {
      const transactionInputs = selectedOccurrenceKeys(source, partial.selected);
      selectedPrepared = finalFrameRow(
        source,
        acquiredRow,
        ownership,
        context.availableHeightPt - availablePt,
        context,
        floatRegistry,
        floatParagraphId,
        rowCursor,
        (occurrence) => transactionInputs.has(occurrenceSelectionKey(occurrence)),
      );
      const reselection = partialRow(
        source, selectedPrepared.row, rowCursor, availablePt, context,
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
    }
    break;
  }

  const sourceRows = selected.filter((row) => row.ownership === 'source');
  if (sourceRows.length === 0) {
    const canProgressOnFreshPage = context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt;
    return {
      fragment: null,
      nextCursor: cursor,
      requiresFreshPage: canProgressOnFreshPage,
    };
  }
  let fragment = materializeFragment(source, selected, context);
  while (fragment.advancePt > context.availableHeightPt + EPSILON_PT) {
    const last = selected.at(-1);
    const sourceCount = selected.filter((row) => row.ownership === 'source').length;
    const wholeSourceRow = last?.ownership === 'source'
      && last.fragmentIndex === 0
      && last.ranges.every((ranges) => ranges.every((range) => range.kind === 'whole'));
    if (!wholeSourceRow || sourceCount <= 1) break;
    selected.pop();
    nextCursor = Object.freeze({
      rowIndex: last.logicalRowIndex,
      rowFragmentIndex: 0,
      cells: Object.freeze([]),
    });
    fragment = materializeFragment(source, selected, context);
  }
  if (fragment.advancePt > context.availableHeightPt + EPSILON_PT
    && context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt
    && fragment.advancePt <= context.freshPageHeightPt + EPSILON_PT) {
    return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
  }
  return {
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
  };
}
