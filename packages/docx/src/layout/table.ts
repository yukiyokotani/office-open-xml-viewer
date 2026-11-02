import { resolveBorderConflict, type BorderCandidate } from '../cell-border-conflict.js';
import { retainedBorderTreatment } from './border-treatment.js';
import { paragraphGapPt } from './paragraph-spacing.js';
import { snapshotPlainData } from './plain-data.js';
import { tableCellHorizontalSpacingInsets } from './table-columns.js';
import { firstAuthoredTableBorder } from './table-border-layer.js';
import { unionLayoutRects } from './rect-union.js';
import {
  wordAlignedTableOriginPt,
  wordCollapsedBorderRowTrackFootprintPt,
  wordExactRowFloorPt,
  wordExactRowVerticalClipBounds,
  wordSpacedCellInsideBorderOverridesTable,
} from './table-compatibility.js';
import type {
  BlockLayoutResult,
  FlowBlockPlacement,
  LayoutRect,
  LayoutServices,
  Matrix2DData,
  ParagraphLayout,
  ResolvedBorderSegment,
  TableBorderInput,
  TableCellBlockInput,
  TableCellBlockLayout,
  TableCellLayout,
  TableCellLayoutInput,
  TableEdgeInputs,
  TableLayout,
  TableLayoutInput,
  TableRowLayout,
  TableRowLayoutInput,
} from './types.js';

interface CellFlowGeometry {
  readonly blocks: readonly TableCellBlockLayout[];
  readonly flowHeightPt: number;
  readonly inkTopPt: number;
  readonly inkHeightPt: number;
  readonly cellContainmentTopPt: number | null;
  readonly cellContainmentBottomPt: number | null;
}

interface CellOwner {
  readonly input: TableCellLayoutInput;
  readonly rowIndex: number;
  readonly lastRowIndex: number;
}

interface ResolvedBoundary {
  readonly border: TableBorderInput;
  readonly edge: ResolvedBorderSegment['edge'];
}

interface HorizontalBoundarySide {
  readonly owner: CellOwner | null;
  readonly border: BorderCandidate | null;
}

interface HorizontalBoundary {
  readonly above: HorizontalBoundarySide;
  readonly below: HorizontalBoundarySide;
  readonly edge: ResolvedBorderSegment['edge'];
}

function paragraphBlockAdvance(layout: ParagraphLayout): number {
  return Math.max(0, layout.advancePt - layout.spacing.beforePt - layout.spacing.afterPt);
}

/**
 * Fold the retained child layouts once. Paragraph spacing is a relationship
 * between adjacent paragraphs, so it cannot be recovered by summing each
 * child's advance independently after table placement.
 */
function resolveCellFlow(blocks: readonly TableCellBlockInput[]): CellFlowGeometry {
  const placements: TableCellBlockLayout[] = [];
  let cursorPt = 0;
  let previousParagraph: ParagraphLayout | null = null;
  let previousAfterPt = 0;
  let firstInkTopPt: number | undefined;
  let lastInkBottomPt = 0;
  let cellContainmentTopPt: number | null = null;
  let cellContainmentBottomPt: number | null = null;

  for (const block of blocks) {
    const layout = block.layout;
    if (layout.kind === 'paragraph') {
      const beforePt = layout.spacing.beforePt;
      const afterPt = layout.spacing.afterPt;
      const gapPt = previousParagraph
        ? paragraphGapPt(previousParagraph, layout, previousAfterPt, beforePt)
        : beforePt;
      const advancePt = block.structuralTrailing ? 0 : paragraphBlockAdvance(layout);
      const offsetPt = cursorPt + (block.structuralTrailing ? 0 : gapPt);
      placements.push({ layout, offsetPt, advancePt });
      if (!block.structuralTrailing) {
        cursorPt = offsetPt + advancePt;
        firstInkTopPt ??= offsetPt;
        lastInkBottomPt = Math.max(lastInkBottomPt, cursorPt);
        previousParagraph = layout;
        previousAfterPt = afterPt;
      }
      if (layout.cellContainmentBounds) {
        const containmentTopPt = offsetPt
          + layout.cellContainmentBounds.yPt
          - layout.flowBounds.yPt;
        const containmentBottomPt = containmentTopPt
          + layout.cellContainmentBounds.heightPt;
        firstInkTopPt = firstInkTopPt === undefined
          ? containmentTopPt
          : Math.min(firstInkTopPt, containmentTopPt);
        lastInkBottomPt = Math.max(lastInkBottomPt, containmentBottomPt);
        cellContainmentTopPt = cellContainmentTopPt === null
          ? containmentTopPt
          : Math.min(cellContainmentTopPt, containmentTopPt);
        cellContainmentBottomPt = cellContainmentBottomPt === null
          ? containmentBottomPt
          : Math.max(cellContainmentBottomPt, containmentBottomPt);
      }
      continue;
    }

    if (previousParagraph) cursorPt += previousAfterPt;
    const advancePt = layout.advancePt;
    placements.push({ layout, offsetPt: cursorPt, advancePt });
    firstInkTopPt ??= cursorPt;
    cursorPt += advancePt;
    lastInkBottomPt = cursorPt;
    previousParagraph = null;
    previousAfterPt = 0;
  }

  const flowHeightPt = cursorPt + (previousParagraph ? previousAfterPt : 0);
  const inkTopPt = firstInkTopPt ?? 0;
  return {
    blocks: placements,
    flowHeightPt,
    inkTopPt,
    inkHeightPt: Math.max(0, lastInkBottomPt - inkTopPt),
    cellContainmentTopPt,
    cellContainmentBottomPt,
  };
}

function cellContentRequiredHeightPt(flow: CellFlowGeometry): number {
  const containmentTopPt = flow.cellContainmentTopPt ?? 0;
  const containmentBottomPt = flow.cellContainmentBottomPt ?? 0;
  // ECMA-376 §20.4.2.3 layoutInCell contains the object frame even though the
  // object remains out of ordinary paragraph flow. General child ink is not a
  // row-sizing input; only this explicitly retained containment interval is.
  return Math.max(flow.flowHeightPt, containmentBottomPt)
    - Math.min(0, containmentTopPt);
}

/**
 * Return the authored block-flow height used by table row layout. Pagination
 * calls this same fold while choosing legal fragment boundaries so paragraph
 * spacing collapse cannot disagree with the geometry materialized afterwards.
 */
export function measureTableCellBlockFlowHeightPt(
  blocks: readonly TableCellBlockInput[],
): number {
  return cellContentRequiredHeightPt(resolveCellFlow(blocks));
}

interface RowSpacingInsets {
  readonly topPt: number;
  readonly bottomPt: number;
}

function effectiveCellSpacingPt(row: TableRowLayoutInput | undefined): number {
  return Number.isFinite(row?.cellSpacingPt) ? Math.max(0, row?.cellSpacingPt ?? 0) : 0;
}

/**
 * §17.4.43/.44/.45 require spacing between adjacent cells and the table edge
 * without increasing the table width. An outer boundary owns the full row
 * spacing outside WORD_TABLE_CELL_SPACING_GEOMETRY; an internal boundary is one shared gap split symmetrically between
 * its two cells. When adjacent rows differ, the larger minimum owns the shared
 * boundary so neither row's constraint is weakened.
 */
function rowSpacingInsets(
  rows: readonly TableRowLayoutInput[],
  rowIndex: number,
): RowSpacingInsets {
  const currentPt = effectiveCellSpacingPt(rows[rowIndex]);
  if (currentPt > 0 && rows[rowIndex]?.spacingGeometry) {
    return { topPt: rowIndex === 0 ? 2 * currentPt : currentPt,
      bottomPt: rowIndex === rows.length - 1 ? 2 * currentPt : currentPt };
  }
  const previousPt = effectiveCellSpacingPt(rows[rowIndex - 1]);
  const nextPt = effectiveCellSpacingPt(rows[rowIndex + 1]);
  return {
    topPt: rowIndex === 0 ? currentPt : Math.max(previousPt, currentPt) / 2,
    bottomPt: rowIndex === rows.length - 1 ? currentPt : Math.max(currentPt, nextPt) / 2,
  };
}

function cellRequiredHeight(
  input: TableCellLayoutInput,
  flow: CellFlowGeometry,
  spacing: RowSpacingInsets,
): number {
  // ECMA-376 §17.4.72: rotated text runs along the cell's height, so the
  // physical height requirement is the acquired line length, not the stacked
  // block extent (which spans the cell width instead).
  const contentPt = input.verticalText
    ? input.verticalText.requiredLineLengthPt
    : cellContentRequiredHeightPt(flow);
  return spacing.topPt
    + input.margins.topPt
    + contentPt
    + input.margins.bottomPt
    + spacing.bottomPt;
}

/** Local-to-table matrix for a rotated cell's content rectangle. The local
 * frame is the horizontal layout frame: x runs along each line, y across
 * lines. `vert`/`eaVert` turn it a quarter clockwise (first line at the right
 * edge, text top to bottom, §17.18.93 tbRl/tbRlV); `vert270` a quarter
 * counter-clockwise (first line at the left edge, text bottom to top, btLr). */
function verticalCellTransform(
  mode: NonNullable<TableCellLayoutInput['verticalText']>['mode'],
  content: LayoutRect,
): Matrix2DData {
  return mode === 'vert270'
    ? { a: 0, b: -1, c: 1, d: 0, e: content.xPt, f: content.yPt + content.heightPt }
    : { a: 0, b: 1, c: -1, d: 0, e: content.xPt + content.widthPt, f: content.yPt };
}

export function mergeEndRow(
  rows: readonly TableRowLayoutInput[],
  startRow: number,
  columnStart: number,
  columnSpan: number,
): number {
  let endRow = startRow;
  for (let rowIndex = startRow + 1; rowIndex < rows.length; rowIndex += 1) {
    const continuation = rows[rowIndex]?.cells.find((cell) => (
      cell.columnStart === columnStart
      && cell.columnSpan === columnSpan
      && cell.verticalMerge === 'continue'
    ));
    if (!continuation) break;
    endRow = rowIndex;
  }
  return endRow;
}

/**
 * The vertical-merge role `cell` of `row` (input row `rowIndex`) plays in
 * this input's grid. ECMA-376 §17.4.84 relates a continuation to the merged
 * cell above it in the same grid, so every cell keeps its authored role,
 * except in a cell-owner segment projection
 * (`segmentOpeningLogicalRowIndex`, table-owner-runs.ts): there a
 * continuation of the segment's first own row with no merged cell above it in
 * this grid lost its restart to another segment, and opens its region as an
 * empty owner with its own borders, shading (materializeTableRow) and margins
 * (resolveRowTrack) (library projected-segment contract). Its content stays
 * suppressed, since its authored role is unchanged. The row is matched by its
 * logical index, so every occurrence of it — each fragment of a row cut
 * across pages included, whose id is its own — is projected alike: a fragment
 * that opens its own grid has no merged cell above it either.
 */
function gridMergeRole(
  input: TableLayoutInput,
  row: TableRowLayoutInput,
  rowIndex: number,
  cell: TableCellLayoutInput,
): TableCellLayoutInput['verticalMerge'] {
  return projectedMergeRole(input.segmentOpeningLogicalRowIndex, input.rows[rowIndex - 1], row, cell);
}

/** {@link gridMergeRole} given the row above `row` in the grid directly, for
 * the row-track solver, which reads windows of the input's rows, and for
 * pagination's bounded track windows (table-pagination.ts
 * completedPartialRowWindowEnd), which must read the role the window's own
 * layout gives each cell. */
export function projectedMergeRole(
  segmentOpeningLogicalRowIndex: TableLayoutInput['segmentOpeningLogicalRowIndex'],
  above: TableRowLayoutInput | undefined,
  row: TableRowLayoutInput,
  cell: TableCellLayoutInput,
): TableCellLayoutInput['verticalMerge'] {
  if (cell.verticalMerge !== 'continue' || row.logicalRowIndex !== segmentOpeningLogicalRowIndex) {
    return cell.verticalMerge;
  }
  const mergedAbove = above?.cells.some((candidate) => (
    candidate.verticalMerge !== 'none'
    && candidate.columnStart === cell.columnStart
    && candidate.columnSpan === cell.columnSpan
  )) ?? false;
  return mergedAbove ? 'continue' : 'restart';
}

function semanticRowFloor(row: TableRowLayoutInput): number {
  if (row.heightRule === 'exact') {
    // `word-exact-row-height-bottom-padding`: compatibility-owned exact-row floor.
    return wordExactRowFloorPt(
      row.heightPt,
      row.cells.map((cell) => cell.margins.bottomPt),
    );
  }
  if (row.heightRule === 'atLeast') return Math.max(0, row.heightPt ?? 0);
  // ECMA-376 §17.4.80: an explicit auto rule has no predetermined minimum.
  // The parser adapter maps Word's omitted-hRule behavior to atLeast before
  // this normalized contract is built.
  return 0;
}

/** A vertical merge whose owner is in a resolved row and whose next row
 * continues it: its constraint is still open. */
interface OpenMergeTrack {
  readonly cell: TableCellLayoutInput;
  readonly start: number;
  readonly topInsetPt: number;
}

function continuesMerge(cell: TableCellLayoutInput, owner: TableCellLayoutInput): boolean {
  // Matched by grid columns, as mergeEndRow matches a continuation.
  return cell.verticalMerge === 'continue'
    && cell.columnStart === owner.columnStart
    && cell.columnSpan === owner.columnSpan;
}

/** Row tracks resolved so far, before collapsed-rule footprints. */
interface RowTrackStore {
  readonly length: number;
  push(heightPt: number, contentHeightPt: number, exact: boolean): void;
  /** The current heights of rows `start..end`, summed from `start`. */
  sum(start: number, end: number): number;
  grow(rowIndex: number, deficitPt: number): void;
  /** The last row at or above `rowIndex` that is not `exact`, or -1. */
  lastNonExact(rowIndex: number): number;
  raiseContentHeight(rowIndex: number, requiredPt: number): void;
}

/**
 * Resolve row `rows[index]` as the next row of `store`, whose following row
 * is `rows[index + 1]` (none: the row is the table's last). `open` holds the
 * merges continuing into it from above, in owner order; the result holds the
 * merges continuing out of it. A merge's constraint is applied when its last
 * row is resolved, so constraints are applied ordered by end row, then start
 * row, then owner order — the order the interval policy below defines — and
 * each reads only rows already resolved. Every row's floor and unmerged cells
 * are therefore fixed before any constraint reaches it, and a resolved row
 * changes later only by a deficit of a merge ending below it.
 *
 * Cells take their grid role (gridMergeRole, `segmentOpeningLogicalRowIndex`
 * naming the input's opening row, `rows[index - 1]` the row above it). A projected empty
 * owner is therefore one more interval constraint: ECMA-376 §17.4.68 tcMar
 * makes its own margins part of its cell, and §17.4.80 lets an auto row grow
 * to what its cells require. Its flow stays empty (`flowOf` keeps the
 * authored continuation's content suppressed), so it requires its margins and
 * row spacing only, with no assumed paragraph-mark height (library
 * projected-segment contract, not a Word measurement).
 */
function resolveRowTrack(
  store: RowTrackStore,
  rows: readonly TableRowLayoutInput[],
  index: number,
  open: readonly OpenMergeTrack[],
  flowOf: (cell: TableCellLayoutInput) => CellFlowGeometry,
  segmentOpeningLogicalRowIndex: TableLayoutInput['segmentOpeningLogicalRowIndex'],
): readonly OpenMergeTrack[] {
  const row = rows[index]!;
  const next = rows[index + 1];
  const rowIndex = store.length;
  const spacing = rowSpacingInsets(rows, index);
  const roles = row.cells.map((cell) => projectedMergeRole(
    segmentOpeningLogicalRowIndex, rows[index - 1], row, cell,
  ));
  let heightPt = semanticRowFloor(row);
  let contentHeightPt = 0;
  for (const [cellIndex, cell] of row.cells.entries()) {
    if (roles[cellIndex] !== 'none') continue;
    const required = cellRequiredHeight(cell, flowOf(cell), spacing)
      + 2 * (row.spacingGeometry?.borderPt ?? 0);
    contentHeightPt = Math.max(contentHeightPt, required);
    if (row.heightRule !== 'exact') heightPt = Math.max(heightPt, required);
  }
  store.push(heightPt, contentHeightPt, row.heightRule === 'exact');
  // A merged owner is one interval constraint over row tracks. ECMA-376 defines
  // the merged region but not deficit distribution. The terminal-growable greedy
  // policy makes the minimum total change, preserves earlier boundaries, reuses
  // prior interval growth, and never violates an exact track.
  const owners: OpenMergeTrack[] = [...open];
  for (const [cellIndex, cell] of row.cells.entries()) {
    if (roles[cellIndex] === 'restart') owners.push({ cell, start: rowIndex, topInsetPt: spacing.topPt });
  }
  const stillOpen: OpenMergeTrack[] = [];
  for (const merge of owners) {
    if (next?.cells.some((cell) => continuesMerge(cell, merge.cell))) {
      stillOpen.push(merge);
      continue;
    }
    const requiredPt = cellRequiredHeight(
      merge.cell,
      flowOf(merge.cell),
      { topPt: merge.topInsetPt, bottomPt: spacing.bottomPt },
    );
    store.raiseContentHeight(merge.start, requiredPt);
    const deficitPt = requiredPt - store.sum(merge.start, rowIndex);
    if (deficitPt <= 0) continue;
    const target = store.lastNonExact(rowIndex);
    if (target >= merge.start) store.grow(target, deficitPt);
  }
  return stillOpen;
}

function resolveRowHeights(
  input: TableLayoutInput,
  flows: ReadonlyMap<string, CellFlowGeometry>,
  boundaryFootprintsPt: readonly number[],
): Readonly<{ heights: readonly number[]; contentHeights: readonly number[] }> {
  const rows = input.rows;
  const heights: number[] = [];
  const contentHeights: number[] = [];
  const lastNonExact: number[] = [];
  const store: RowTrackStore = {
    get length() { return heights.length; },
    push(heightPt, contentHeightPt, exact) {
      lastNonExact.push(exact ? (lastNonExact.at(-1) ?? -1) : heights.length);
      heights.push(heightPt);
      contentHeights.push(contentHeightPt);
    },
    sum(start, end) {
      let currentPt = 0;
      for (let rowIndex = start; rowIndex <= end; rowIndex += 1) currentPt += heights[rowIndex] ?? 0;
      return currentPt;
    },
    grow(rowIndex, deficitPt) { heights[rowIndex] = (heights[rowIndex] ?? 0) + deficitPt; },
    lastNonExact: (rowIndex) => lastNonExact[rowIndex] ?? -1,
    raiseContentHeight(rowIndex, requiredPt) {
      contentHeights[rowIndex] = Math.max(contentHeights[rowIndex] ?? 0, requiredPt);
    },
  };
  const flowOf = (cell: TableCellLayoutInput) => flows.get(cell.id) ?? resolveCellFlow([]);
  let open: readonly OpenMergeTrack[] = [];
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    open = resolveRowTrack(store, rows, rowIndex, open, flowOf, input.segmentOpeningLogicalRowIndex);
  }
  rows.forEach((_row, rowIndex) => {
    heights[rowIndex] = (heights[rowIndex] ?? 0) + (boundaryFootprintsPt[rowIndex] ?? 0);
  });
  return { heights, contentHeights };
}

function toConflictCandidate(
  border: TableBorderInput | null,
  source: BorderCandidate['source'],
): BorderCandidate | null {
  if (!border) return null;
  return {
    source,
    spec: {
      width: border.widthPt,
      color: border.color,
      style: border.authoredStyle,
    },
  };
}

function physicalCellEdges(
  cell: TableCellLayoutInput,
  table: TableEdgeInputs,
  exception: TableEdgeInputs | null,
  rowIndex: number,
  lastRowIndex: number,
  rowCount: number,
  columnCount: number,
  bidiVisual: boolean,
): Readonly<{
  top: BorderCandidate | null;
  right: BorderCandidate | null;
  bottom: BorderCandidate | null;
  left: BorderCandidate | null;
}> {
  const cascade = (
    direct: TableBorderInput | null,
    cellInside: TableBorderInput | null,
    exceptionOuter: TableBorderInput | null,
    exceptionInside: TableBorderInput | null,
    tableOuter: TableBorderInput | null,
    tableInside: TableBorderInput | null,
    useInside: boolean,
  ): BorderCandidate | null => {
    const resolvedCell = firstAuthoredTableBorder(direct, useInside ? cellInside : null);
    if (resolvedCell) return toConflictCandidate(resolvedCell, 'cell');
    return toConflictCandidate(
      useInside
        ? firstAuthoredTableBorder(exceptionInside, tableInside)
        : firstAuthoredTableBorder(exceptionOuter, tableOuter),
      'table',
    );
  };
  const top = cascade(
    cell.borders.top,
    cell.borders.insideH,
    exception?.top ?? null,
    exception?.insideH ?? null,
    table.top,
    table.insideH,
    rowIndex !== 0,
  );
  const bottom = cascade(
    cell.borders.bottom,
    cell.borders.insideH,
    exception?.bottom ?? null,
    exception?.insideH ?? null,
    table.bottom,
    table.insideH,
    lastRowIndex !== rowCount - 1,
  );
  const logicalLeft = cascade(
    cell.borders.left,
    cell.borders.insideV,
    exception?.left ?? null,
    exception?.insideV ?? null,
    table.left,
    table.insideV,
    cell.columnStart !== 0,
  );
  const logicalRight = cascade(
    cell.borders.right,
    cell.borders.insideV,
    exception?.right ?? null,
    exception?.insideV ?? null,
    table.right,
    table.insideV,
    cell.columnStart + cell.columnSpan !== columnCount,
  );
  return bidiVisual
    ? { top, right: logicalLeft, bottom, left: logicalRight }
    : { top, right: logicalRight, bottom, left: logicalLeft };
}

function candidateInput(candidate: BorderCandidate | null): TableBorderInput | null {
  if (!candidate) return null;
  return {
    widthPt: candidate.spec.width,
    color: candidate.spec.color ?? '#000000',
    authoredStyle: candidate.spec.style,
  };
}

function resolveBoundary(
  first: BorderCandidate | null,
  second: BorderCandidate | null,
  edge: ResolvedBorderSegment['edge'],
): ResolvedBoundary | null {
  const winner = candidateInput(resolveBorderConflict(first, second));
  return winner ? { border: winner, edge } : null;
}

function ownerGrid(
  input: TableLayoutInput,
): Readonly<{
  owners: readonly CellOwner[];
  occupancy: readonly (readonly number[])[];
}> {
  const columnCount = input.columnWidthsPt.length;
  const owners: CellOwner[] = [];
  const occupancy = input.rows.map(() => new Array<number>(columnCount).fill(-1));
  input.rows.forEach((row, rowIndex) => {
    for (const cell of row.cells) {
      const role = gridMergeRole(input, row, rowIndex, cell);
      if (role === 'continue') continue;
      const lastRowIndex = role === 'restart'
        ? mergeEndRow(input.rows, rowIndex, cell.columnStart, cell.columnSpan)
        : rowIndex;
      const ownerIndex = owners.length;
      owners.push({ input: cell, rowIndex, lastRowIndex });
      const endColumn = Math.min(columnCount, cell.columnStart + cell.columnSpan);
      for (let coveredRow = rowIndex; coveredRow <= lastRowIndex; coveredRow += 1) {
        for (let column = Math.max(0, cell.columnStart); column < endColumn; column += 1) {
          occupancy[coveredRow]![column] = ownerIndex;
        }
      }
    }
  });
  return { owners, occupancy };
}

// `word-vertical-merge-terminal-border`: the terminal continuation cell owns the bottom edge.
function terminalMergeCell(
  input: TableLayoutInput,
  owner: CellOwner,
): TableCellLayoutInput {
  if (owner.lastRowIndex === owner.rowIndex) return owner.input;
  return input.rows[owner.lastRowIndex]?.cells.find((cell) => (
    cell.columnStart === owner.input.columnStart
    && cell.columnSpan === owner.input.columnSpan
    && cell.verticalMerge === 'continue'
  )) ?? owner.input;
}

function resolvedBoundaries(input: TableLayoutInput): Readonly<{
  horizontal: readonly (readonly (HorizontalBoundary | null)[])[];
  vertical: readonly (readonly (ResolvedBoundary | null)[])[];
  occupancy: readonly (readonly number[])[];
}> {
  const rowCount = input.rows.length;
  const columnCount = input.columnWidthsPt.length;
  const { owners, occupancy } = ownerGrid(input);
  const edgesOf = (ownerIndex: number, terminal = false) => {
    const owner = owners[ownerIndex];
    if (!owner) return null;
    const cell = terminal ? terminalMergeCell(input, owner) : owner.input;
    const exceptionRowIndex = terminal && cell !== owner.input
      ? owner.lastRowIndex
      : owner.rowIndex;
    return physicalCellEdges(
      cell,
      input.borders,
      input.rows[exceptionRowIndex]?.exceptionBorders ?? null,
      owner.rowIndex,
      owner.lastRowIndex,
      rowCount,
      columnCount,
      input.bidiVisual,
    );
  };

  const horizontal = Array.from({ length: rowCount + 1 }, (_unused, boundary) => (
    Array.from({ length: columnCount }, (_cell, column) => {
      const aboveIndex = boundary > 0 ? occupancy[boundary - 1]?.[column] ?? -1 : -1;
      const belowIndex = boundary < rowCount ? occupancy[boundary]?.[column] ?? -1 : -1;
      if (aboveIndex >= 0 && aboveIndex === belowIndex) return null;
      const below = edgesOf(belowIndex);
      const edge: HorizontalBoundary['edge'] = boundary === 0
        ? 'top'
        : boundary === rowCount ? 'bottom' : 'between';
      return {
        above: {
          owner: owners[aboveIndex] ?? null,
          border: edgesOf(aboveIndex, true)?.bottom ?? null,
        },
        below: {
          owner: owners[belowIndex] ?? null,
          border: below?.top ?? null,
        },
        edge,
      };
    })
  ));

  const vertical = Array.from({ length: columnCount + 1 }, (_unused, boundary) => (
    Array.from({ length: rowCount }, (_row, rowIndex) => {
      const logicalBefore = boundary > 0 ? occupancy[rowIndex]?.[boundary - 1] ?? -1 : -1;
      const logicalAfter = boundary < columnCount ? occupancy[rowIndex]?.[boundary] ?? -1 : -1;
      const physicalLeftIndex = input.bidiVisual ? logicalAfter : logicalBefore;
      const physicalRightIndex = input.bidiVisual ? logicalBefore : logicalAfter;
      if (physicalLeftIndex >= 0 && physicalLeftIndex === physicalRightIndex) return null;
      return resolveBoundary(
        edgesOf(physicalLeftIndex)?.right ?? null,
        edgesOf(physicalRightIndex)?.left ?? null,
        boundary === 0
          ? (input.bidiVisual ? 'right' : 'left')
          : boundary === columnCount
            ? (input.bidiVisual ? 'left' : 'right')
            : 'between',
      );
    })
  ));
  return { horizontal, vertical, occupancy };
}

/** Winning collapsed rules contribute their page-local half-rule footprint to
 * non-exact Word row tracks. Spaced cells keep independent edge boxes. */
function collapsedHorizontalBoundaryWidthsPt(
  input: TableLayoutInput,
  boundaries: ReturnType<typeof resolvedBoundaries>,
): readonly number[] {
  return boundaries.horizontal.map((segments, boundaryIndex) => {
    if (
      effectiveCellSpacingPt(input.rows[boundaryIndex - 1]) > 0
      || effectiveCellSpacingPt(input.rows[boundaryIndex]) > 0
    ) return 0;
    return segments.reduce((maximumPt, segment) => {
      if (!segment) return maximumPt;
      const winner = resolveBoundary(segment.above.border, segment.below.border, segment.edge);
      return Math.max(maximumPt, winner?.border.widthPt ?? 0);
    }, 0);
  });
}

/** The collapsed horizontal rules each row's content sits inside
 * (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT): the full rule above the row
 * and, for the last row, the full bottom rule below it. */
export interface RowRuleInsets {
  readonly top: readonly number[];
  readonly bottom: readonly number[];
}

function rowRuleInsetsPt(
  input: TableLayoutInput,
  boundaries: ReturnType<typeof resolvedBoundaries>,
): RowRuleInsets {
  const separated = input.rows[0]?.spacingGeometry;
  if (separated && effectiveCellSpacingPt(input.rows[0]) > 0) {
    return { top: input.rows.map((_row, index) => index === 0 ? separated.borderPt : 0),
      bottom: input.rows.map((_row, index) => index === input.rows.length - 1 ? separated.borderPt : 0) };
  }
  const widthsPt = collapsedHorizontalBoundaryWidthsPt(input, boundaries);
  const last = input.rows.length - 1;
  return {
    top: input.rows.map((_row, rowIndex) => widthsPt[rowIndex] ?? 0),
    bottom: input.rows.map((_row, rowIndex) => rowIndex === last ? widthsPt[rowIndex + 1] ?? 0 : 0),
  };
}

function rowBoundaryFootprintsPt(
  input: TableLayoutInput,
  boundaries: ReturnType<typeof resolvedBoundaries>,
  insets: RowRuleInsets = rowRuleInsetsPt(input, boundaries),
): readonly number[] {
  return input.rows.map((row, rowIndex) => wordCollapsedBorderRowTrackFootprintPt(
    row.heightRule === 'exact',
    insets.top[rowIndex] ?? 0,
    insets.bottom[rowIndex] ?? 0,
  ));
}

/** The collapsed rule insets of each row (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT). */
export function tableRowRuleInsetsPt(input: TableLayoutInput): RowRuleInsets {
  return rowRuleInsetsPt(input, resolvedBoundaries(input));
}

/** Resolve the page-local collapsed-rule footprint charged to each row track.
 * Pagination uses the same layout authority before selecting partial content,
 * so fragment materialization cannot introduce an unreserved border advance. */
export function tableRowBoundaryFootprintsPt(input: TableLayoutInput): readonly number[] {
  return rowBoundaryFootprintsPt(input, resolvedBoundaries(input));
}

function borderSegment(
  resolved: ResolvedBoundary,
  from: Readonly<{ xPt: number; yPt: number }>,
  to: Readonly<{ xPt: number; yPt: number }>,
): ResolvedBorderSegment {
  return {
    ...(resolved.edge ? { edge: resolved.edge } : {}),
    from,
    to,
    color: resolved.border.color,
    widthPt: resolved.border.widthPt,
    ...retainedBorderTreatment(resolved.border.authoredStyle, resolved.border.widthPt),
  };
}

const noTableEdges: TableEdgeInputs = Object.freeze({
  top: null,
  right: null,
  bottom: null,
  left: null,
  insideH: null,
  insideV: null,
});

function visibleBorder(candidate: BorderCandidate | null): TableBorderInput | null {
  const border = candidateInput(candidate);
  return border && border.authoredStyle !== 'nil' && border.authoredStyle !== 'none'
    ? border
    : null;
}

function materializeBorders(
  input: TableLayoutInput,
  rowXPt: readonly number[],
  tableYPt: number,
  rowHeightsPt: readonly number[],
  boundaries: ReturnType<typeof resolvedBoundaries>,
): readonly ResolvedBorderSegment[] {
  const columnOffsets = [0];
  for (const width of input.columnWidthsPt) {
    columnOffsets.push((columnOffsets.at(-1) ?? 0) + width);
  }
  const rowOffsets = [0];
  for (const height of rowHeightsPt) rowOffsets.push((rowOffsets.at(-1) ?? 0) + height);
  const tableWidthPt = columnOffsets.at(-1) ?? 0;
  const columnX = (rowIndex: number, column: number) => (rowXPt[rowIndex] ?? 0) + (
    input.bidiVisual ? tableWidthPt - (columnOffsets[column] ?? 0) : (columnOffsets[column] ?? 0)
  );
  const rowY = (row: number) => tableYPt + (rowOffsets[row] ?? 0);
  const segments: ResolvedBorderSegment[] = [];

  const push = (
    border: TableBorderInput | null,
    edge: ResolvedBorderSegment['edge'],
    from: Readonly<{ xPt: number; yPt: number }>,
    to: Readonly<{ xPt: number; yPt: number }>,
  ): void => {
    if (!border || border.authoredStyle === 'nil' || border.authoredStyle === 'none') return;
    segments.push(borderSegment({ border, edge }, from, to));
  };

  // A spanning/merged owner occupies several logical slots but paints one
  // detached top or bottom edge at a spaced boundary.
  const detachedHorizontalOwners = new Set<string>();
  const pushDetachedHorizontal = (
    side: HorizontalBoundarySide,
    boundary: number,
    position: 'top' | 'bottom',
    edge: ResolvedBorderSegment['edge'],
  ): void => {
    const owner = side.owner;
    if (!owner) return;
    const key = `${boundary}:${position}:${owner.input.id}`;
    if (detachedHorizontalOwners.has(key)) return;
    detachedHorizontalOwners.add(key);
    const row = input.rows[owner.rowIndex];
    if (!row) return;
    const spacingPt = effectiveCellSpacingPt(row);
    const startXPt = columnX(owner.rowIndex, owner.input.columnStart);
    const endXPt = columnX(owner.rowIndex, Math.min(
      input.columnWidthsPt.length,
      owner.input.columnStart + owner.input.columnSpan,
    ));
    const { startPt: logicalStartInsetPt, endPt: logicalEndInsetPt } =
      tableCellHorizontalSpacingInsets(
        spacingPt,
        owner.input.columnStart,
        owner.input.columnSpan,
        input.columnWidthsPt.length,
      );
    const leftPt = Math.min(startXPt, endXPt)
      + (input.bidiVisual ? logicalEndInsetPt : logicalStartInsetPt);
    const rightPt = Math.max(startXPt, endXPt)
      - (input.bidiVisual ? logicalStartInsetPt : logicalEndInsetPt);
    const topPt = rowY(owner.rowIndex) + rowSpacingInsets(input.rows, owner.rowIndex).topPt;
    const bottomPt = rowY(owner.lastRowIndex + 1)
      - rowSpacingInsets(input.rows, owner.lastRowIndex).bottomPt;
    const edges = physicalCellEdges(
      owner.input,
      noTableEdges,
      null,
      owner.rowIndex,
      owner.lastRowIndex,
      input.rows.length,
      input.columnWidthsPt.length,
      input.bidiVisual,
    );
    const candidate = position === 'top' ? edges.top : edges.bottom;
    const yPt = position === 'top' ? topPt : bottomPt;
    push(visibleBorder(candidate), edge, { xPt: leftPt, yPt }, { xPt: rightPt, yPt });
  };

  // WORD_TABLE_CELL_SPACING_GEOMETRY: table and cells own independent
  // border boxes; there is no extra inside rule down the middle of a gap.
  const separated = input.rows[0]?.spacingGeometry;
  if (separated && effectiveCellSpacingPt(input.rows[0]) > 0) {
    const border = input.borders.top;
    const t = separated.borderPt;
    const rectangle = (left: number, top: number, right: number, bottom: number): void => {
      push(border, 'top', { xPt: left, yPt: top + t / 2 }, { xPt: right, yPt: top + t / 2 });
      push(border, 'bottom', { xPt: left, yPt: bottom - t / 2 }, { xPt: right, yPt: bottom - t / 2 });
      push(border, 'left', { xPt: left + t / 2, yPt: top + t / 2 }, { xPt: left + t / 2, yPt: bottom - t / 2 });
      push(border, 'right', { xPt: right - t / 2, yPt: top + t / 2 }, { xPt: right - t / 2, yPt: bottom - t / 2 });
    };
    rectangle(rowXPt[0] ?? 0, tableYPt, (rowXPt[0] ?? 0) + tableWidthPt, rowY(input.rows.length));
    input.rows.forEach((row, rowIndex) => {
      const spacing = rowSpacingInsets(input.rows, rowIndex);
      for (const cell of row.cells) {
        const start = columnX(rowIndex, cell.columnStart);
        const end = columnX(rowIndex, cell.columnStart + cell.columnSpan);
        const insets = tableCellHorizontalSpacingInsets(row.cellSpacingPt, cell.columnStart,
          cell.columnSpan, input.columnWidthsPt.length, true);
        rectangle(Math.min(start, end) + (input.bidiVisual ? insets.endPt : insets.startPt),
          rowY(rowIndex) + spacing.topPt + (rowIndex === 0 ? t : 0),
          Math.max(start, end) - (input.bidiVisual ? insets.startPt : insets.endPt),
          rowY(rowIndex + 1) - spacing.bottomPt - (rowIndex === input.rows.length - 1 ? t : 0));
      }
    });
    return segments;
  }

  boundaries.horizontal.forEach((columns, boundary) => {
    const aboveSpaced = boundary > 0
      && effectiveCellSpacingPt(input.rows[boundary - 1]) > 0;
    const belowSpaced = boundary < input.rows.length
      && effectiveCellSpacingPt(input.rows[boundary]) > 0;
    if (aboveSpaced || belowSpaced) {
      const boundarySpacingPt = Math.max(
        effectiveCellSpacingPt(input.rows[boundary - 1]),
        effectiveCellSpacingPt(input.rows[boundary]),
      );
      const gridRow = belowSpaced ? boundary : boundary - 1;
      const tableXPt = rowXPt[gridRow] ?? 0;
      const edge = boundary === 0
        ? 'top'
        : boundary === input.rows.length ? 'bottom' : 'between';
      if (boundary === 0 || boundary === input.rows.length) {
        const exceptionBorder = boundary === 0
          ? input.rows[0]?.exceptionBorders?.top ?? null
          : input.rows.at(-1)?.exceptionBorders?.bottom ?? null;
        const tableBorder = firstAuthoredTableBorder(
          exceptionBorder,
          boundary === 0 ? input.borders.top : input.borders.bottom,
        );
        push(tableBorder, edge, { xPt: tableXPt, yPt: rowY(boundary) }, {
          xPt: tableXPt + tableWidthPt,
          yPt: rowY(boundary),
        });
      } else {
        columns.forEach((horizontal, column) => {
          const aboveOwner = boundaries.occupancy[boundary - 1]?.[column] ?? -1;
          const belowOwner = boundaries.occupancy[boundary]?.[column] ?? -1;
          const separatesOwners = aboveOwner !== belowOwner
            && (aboveOwner >= 0 || belowOwner >= 0);
          if (!horizontal || !separatesOwners) return;
          const conditionalInsideOverridesTable = [
            { side: horizontal.above, directEdge: 'bottom' as const },
            { side: horizontal.below, directEdge: 'top' as const },
          ].some(({ side, directEdge }) => {
              const owner = side.owner;
              if (!owner) return false;
              // `word-spaced-cell-inside-border-conflict`
              return wordSpacedCellInsideBorderOverridesTable({
                spacingPt: boundarySpacingPt,
                directStyle: owner.input.borders[directEdge]?.authoredStyle,
                conditionalInsideStyle: owner.input.borders.insideH?.authoredStyle,
              });
            });
          if (conditionalInsideOverridesTable) return;
          const startXPt = columnX(gridRow, column);
          const endXPt = columnX(gridRow, column + 1);
          const aboveTableBorder = firstAuthoredTableBorder(
            input.rows[boundary - 1]?.exceptionBorders?.insideH ?? null,
            input.borders.insideH,
          );
          const belowTableBorder = firstAuthoredTableBorder(
            input.rows[boundary]?.exceptionBorders?.insideH ?? null,
            input.borders.insideH,
          );
          const tableBorder = resolveBoundary(
            toConflictCandidate(aboveTableBorder, 'table'),
            toConflictCandidate(belowTableBorder, 'table'),
            edge,
          )?.border ?? null;
          push(tableBorder, edge,
            { xPt: Math.min(startXPt, endXPt), yPt: rowY(boundary) },
            { xPt: Math.max(startXPt, endXPt), yPt: rowY(boundary) });
        });
      }
      columns.forEach((horizontal) => {
        if (!horizontal) return;
        pushDetachedHorizontal(horizontal.above, boundary, 'bottom', horizontal.edge);
        pushDetachedHorizontal(horizontal.below, boundary, 'top', horizontal.edge);
      });
      return;
    }

    const horizontalSegments: Array<{
      resolved: ResolvedBoundary;
      leftPt: number;
      rightPt: number;
    }> = [];
    // Row jc/tblInd can shift tracks far enough that an edge in logical column N
    // overlaps column M in the adjacent row. Flatten owner edges across the whole
    // boundary before sweeping physical X intervals; a per-column conflict pass
    // would double-paint such cross-column overlap.
    const physicalIntervals = new Map<string, Readonly<{
      side: 'above' | 'below';
      border: BorderCandidate;
      leftPt: number;
      rightPt: number;
    }>>();
    columns.forEach((horizontal) => {
      if (!horizontal) return;
      const retainInterval = (
        sideName: 'above' | 'below',
        side: HorizontalBoundarySide,
      ): void => {
        if (!side.owner || !side.border) return;
        const key = `${sideName}:${side.owner.input.id}`;
        if (physicalIntervals.has(key)) return;
        const startPt = columnX(side.owner.rowIndex, side.owner.input.columnStart);
        const endPt = columnX(side.owner.rowIndex, Math.min(
          input.columnWidthsPt.length,
          side.owner.input.columnStart + side.owner.input.columnSpan,
        ));
        physicalIntervals.set(key, {
          side: sideName,
          border: side.border,
          leftPt: Math.min(startPt, endPt),
          rightPt: Math.max(startPt, endPt),
        });
      };
      retainInterval('above', horizontal.above);
      retainInterval('below', horizontal.below);
    });
    const intervals = [...physicalIntervals.values()];
    const breakpoints = [...new Set(intervals.flatMap((interval) => [
      interval.leftPt,
      interval.rightPt,
    ]))].sort((left, right) => left - right);
    const edge: ResolvedBorderSegment['edge'] = boundary === 0
      ? 'top'
      : boundary === input.rows.length ? 'bottom' : 'between';
    for (let index = 1; index < breakpoints.length; index += 1) {
      const leftPt = breakpoints[index - 1] ?? 0;
      const rightPt = breakpoints[index] ?? leftPt;
      if (rightPt <= leftPt) continue;
      const middlePt = (leftPt + rightPt) / 2;
      const active = intervals.filter((interval) => (
        middlePt > interval.leftPt && middlePt < interval.rightPt
      ));
      const above = active.find((interval) => interval.side === 'above')?.border ?? null;
      const below = active.find((interval) => interval.side === 'below')?.border ?? null;
      const resolved = resolveBoundary(above, below, edge);
      if (resolved) horizontalSegments.push({ resolved, leftPt, rightPt });
    }
    horizontalSegments.sort((left, right) => left.leftPt - right.leftPt);
    const merged: typeof horizontalSegments = [];
    for (const segment of horizontalSegments) {
      const previous = merged.at(-1);
      if (previous
        && previous.rightPt === segment.leftPt
        && previous.resolved.edge === segment.resolved.edge
        && previous.resolved.border.widthPt === segment.resolved.border.widthPt
        && previous.resolved.border.color === segment.resolved.border.color
        && previous.resolved.border.authoredStyle === segment.resolved.border.authoredStyle) {
        previous.rightPt = segment.rightPt;
      } else {
        merged.push({ ...segment });
      }
    }
    // WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT: a rule occupies the band
    // below its boundary; the table's bottom rule, the band above its end.
    for (const segment of merged) {
      const halfPt = segment.resolved.border.widthPt / 2;
      const yPt = boundary === input.rows.length ? rowY(boundary) - halfPt : rowY(boundary) + halfPt;
      segments.push(borderSegment(
        segment.resolved,
        { xPt: segment.leftPt, yPt },
        { xPt: segment.rightPt, yPt },
      ));
    }
  });
  // Vertical rules meet the adjacent top and bottom rules at their centre
  // lines, which sit inside the table box (see the horizontal rules above).
  const lastRow = input.rows.length - 1;
  const adjacentRuleWidthPt = (horizontalBoundary: number, columnBoundary: number): number =>
    [columnBoundary - 1, columnBoundary].reduce((maximumPt, column) => {
      const segment = boundaries.horizontal[horizontalBoundary]?.[column];
      if (!segment) return maximumPt;
      const winner = resolveBoundary(segment.above.border, segment.below.border, segment.edge);
      return Math.max(maximumPt, winner?.border.widthPt ?? 0);
    }, 0);
  boundaries.vertical.forEach((rows, boundary) => {
    rows.forEach((resolved, row) => {
      if (effectiveCellSpacingPt(input.rows[row]) > 0) return;
      if (!resolved) return;
      const topPt = rowY(row) + (row === 0 ? adjacentRuleWidthPt(0, boundary) / 2 : 0);
      const bottomPt = rowY(row + 1)
        - (row === lastRow ? adjacentRuleWidthPt(lastRow + 1, boundary) / 2 : 0);
      segments.push(borderSegment(
        resolved,
        { xPt: columnX(row, boundary), yPt: topPt },
        { xPt: columnX(row, boundary), yPt: bottomPt },
      ));
    });
  });

  // ECMA-376 spacing separates opposing cell edge boxes. The narrow
  // compatibility decision below only chooses conditional inside-border
  // winners against the corresponding table inside border.
  input.rows.forEach((row, rowIndex) => {
    const spacingPt = effectiveCellSpacingPt(row);
    if (spacingPt <= 0) return;
    const rowTopPt = rowY(rowIndex);
    const rowBottomPt = rowY(rowIndex + 1);
    const tableXPt = rowXPt[rowIndex] ?? 0;
    push(firstAuthoredTableBorder(row.exceptionBorders?.left ?? null, input.borders.left),
      'left', { xPt: tableXPt, yPt: rowTopPt }, {
      xPt: tableXPt, yPt: rowBottomPt,
    });
    push(firstAuthoredTableBorder(row.exceptionBorders?.right ?? null, input.borders.right),
      'right', { xPt: tableXPt + tableWidthPt, yPt: rowTopPt }, {
      xPt: tableXPt + tableWidthPt, yPt: rowBottomPt,
    });
    const conditionalInsideVBoundaries = new Set<number>();
    for (const cell of row.cells) {
      if (wordSpacedCellInsideBorderOverridesTable({
        spacingPt,
        directStyle: cell.borders.left?.authoredStyle,
        conditionalInsideStyle: cell.borders.insideV?.authoredStyle,
      })) {
        conditionalInsideVBoundaries.add(cell.columnStart);
      }
      if (wordSpacedCellInsideBorderOverridesTable({
        spacingPt,
        directStyle: cell.borders.right?.authoredStyle,
        conditionalInsideStyle: cell.borders.insideV?.authoredStyle,
      })) {
        conditionalInsideVBoundaries.add(cell.columnStart + cell.columnSpan);
      }
    }
    for (let boundary = 1; boundary < input.columnWidthsPt.length; boundary += 1) {
      const logicalBeforeOwner = boundaries.occupancy[rowIndex]?.[boundary - 1] ?? -1;
      const logicalAfterOwner = boundaries.occupancy[rowIndex]?.[boundary] ?? -1;
      const separatesOwners = logicalBeforeOwner !== logicalAfterOwner
        && (logicalBeforeOwner >= 0 || logicalAfterOwner >= 0);
      if (!separatesOwners) continue;
      const xPt = columnX(rowIndex, boundary);
      if (!conditionalInsideVBoundaries.has(boundary)) {
        push(firstAuthoredTableBorder(
          row.exceptionBorders?.insideV ?? null,
          input.borders.insideV,
        ), 'between',
          { xPt, yPt: rowTopPt }, { xPt, yPt: rowBottomPt });
      }
    }

    for (const cell of row.cells) {
      const role = gridMergeRole(input, row, rowIndex, cell);
      if (role === 'continue') continue;
      const lastRowIndex = role === 'restart'
        ? mergeEndRow(input.rows, rowIndex, cell.columnStart, cell.columnSpan)
        : rowIndex;
      const startXPt = columnX(rowIndex, cell.columnStart);
      const endXPt = columnX(rowIndex, Math.min(
        input.columnWidthsPt.length,
        cell.columnStart + cell.columnSpan,
      ));
      const { startPt: logicalStartInsetPt, endPt: logicalEndInsetPt } =
        tableCellHorizontalSpacingInsets(
          spacingPt,
          cell.columnStart,
          cell.columnSpan,
          input.columnWidthsPt.length,
        );
      const leftPt = Math.min(startXPt, endXPt)
        + (input.bidiVisual ? logicalEndInsetPt : logicalStartInsetPt);
      const rightPt = Math.max(startXPt, endXPt)
        - (input.bidiVisual ? logicalStartInsetPt : logicalEndInsetPt);
      const topPt = rowY(rowIndex) + rowSpacingInsets(input.rows, rowIndex).topPt;
      const bottomPt = rowY(lastRowIndex + 1)
        - rowSpacingInsets(input.rows, lastRowIndex).bottomPt;
      const edges = physicalCellEdges(
        cell,
        noTableEdges,
        null,
        rowIndex,
        lastRowIndex,
        input.rows.length,
        input.columnWidthsPt.length,
        input.bidiVisual,
      );
      push(visibleBorder(edges.right), 'right', { xPt: rightPt, yPt: topPt }, { xPt: rightPt, yPt: bottomPt });
      push(visibleBorder(edges.left), 'left', { xPt: leftPt, yPt: topPt }, { xPt: leftPt, yPt: bottomPt });
    }
  });

  // ECMA-376 §17.4.73 tl2br / §17.4.79 tr2bl: a cell diagonal runs between
  // the physical corners of the cell box (the merged box for a vertical or
  // horizontal merge). Diagonals take no part in the §17.4.66 edge conflict
  // resolution, so they are materialized from the owning cell alone.
  input.rows.forEach((row, rowIndex) => {
    for (const cell of row.cells) {
      if (!cell.diagonalBorders) continue;
      const role = gridMergeRole(input, row, rowIndex, cell);
      if (role === 'continue') continue;
      const lastRowIndex = role === 'restart'
        ? mergeEndRow(input.rows, rowIndex, cell.columnStart, cell.columnSpan)
        : rowIndex;
      const startXPt = columnX(rowIndex, cell.columnStart);
      const endXPt = columnX(rowIndex, Math.min(
        input.columnWidthsPt.length,
        cell.columnStart + cell.columnSpan,
      ));
      const { startPt: logicalStartInsetPt, endPt: logicalEndInsetPt } =
        tableCellHorizontalSpacingInsets(
          effectiveCellSpacingPt(row),
          cell.columnStart,
          cell.columnSpan,
          input.columnWidthsPt.length,
        );
      const leftPt = Math.min(startXPt, endXPt)
        + (input.bidiVisual ? logicalEndInsetPt : logicalStartInsetPt);
      const rightPt = Math.max(startXPt, endXPt)
        - (input.bidiVisual ? logicalStartInsetPt : logicalEndInsetPt);
      const topPt = rowY(rowIndex) + rowSpacingInsets(input.rows, rowIndex).topPt;
      const bottomPt = rowY(lastRowIndex + 1)
        - rowSpacingInsets(input.rows, lastRowIndex).bottomPt;
      const { tl2br, tr2bl } = cell.diagonalBorders;
      push(tl2br, undefined, { xPt: leftPt, yPt: topPt }, { xPt: rightPt, yPt: bottomPt });
      push(tr2bl, undefined, { xPt: rightPt, yPt: topPt }, { xPt: leftPt, yPt: bottomPt });
    }
  });
  return segments;
}

function alignedTableOriginX(
  alignment: TableLayoutInput['alignment'],
  indentPt: number,
  bidiVisual: boolean,
  placement: FlowBlockPlacement,
  widthPt: number,
): number {
  const bounds = placement.availableBounds;
  const aligned = alignment === 'center'
    ? bounds.xPt + (bounds.widthPt - widthPt) / 2
    : alignment === 'right'
      ? bounds.xPt + bounds.widthPt - widthPt
      : bounds.xPt;
  if (indentPt === 0) return aligned;
  // Compatibility-owned signed leading-edge translation; bidi reverses the axis.
  return wordAlignedTableOriginPt(aligned, indentPt, bidiVisual);
}

function unionInkBounds(flowBounds: LayoutRect, borders: readonly ResolvedBorderSegment[]): LayoutRect {
  if (borders.length === 0) return flowBounds;
  const left = Math.min(flowBounds.xPt, ...borders.map((item) => Math.min(item.from.xPt, item.to.xPt) - item.widthPt / 2));
  const top = Math.min(flowBounds.yPt, ...borders.map((item) => Math.min(item.from.yPt, item.to.yPt) - item.widthPt / 2));
  const right = Math.max(
    flowBounds.xPt + flowBounds.widthPt,
    ...borders.map((item) => Math.max(item.from.xPt, item.to.xPt) + item.widthPt / 2),
  );
  const bottom = Math.max(
    flowBounds.yPt + flowBounds.heightPt,
    ...borders.map((item) => Math.max(item.from.yPt, item.to.yPt) + item.widthPt / 2),
  );
  return { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
}

/** Recognize complete compound outer frames while retained geometry is still
 * being built. Paint must not infer table structure from independent segments. */
function compoundBorderFrames(
  borders: readonly ResolvedBorderSegment[],
): NonNullable<TableLayout['compoundBorderFrames']> {
  const groups = new Map<string, Array<Readonly<{
    border: ResolvedBorderSegment;
    index: number;
  }>>>();
  borders.forEach((border, index) => {
    if (border.style !== 'compound' || !border.edge || border.edge === 'between') return;
    const key = `${border.authoredStyle}\u0000${border.color}\u0000${border.widthPt}`;
    const group = groups.get(key) ?? [];
    group.push({ border, index });
    groups.set(key, group);
  });
  const frames: NonNullable<TableLayout['compoundBorderFrames']>[number][] = [];
  for (const group of groups.values()) {
    const onEdge = (edge: ResolvedBorderSegment['edge']) =>
      group.filter((item) => item.border.edge === edge);
    const top = onEdge('top');
    const right = onEdge('right');
    const bottom = onEdge('bottom');
    const left = onEdge('left');
    if (!top.length || !right.length || !bottom.length || !left.length) continue;
    const leftPt = Math.min(...top.flatMap(({ border }) => [border.from.xPt, border.to.xPt]));
    const rightPt = Math.max(...top.flatMap(({ border }) => [border.from.xPt, border.to.xPt]));
    const topPt = top[0]!.border.from.yPt;
    const bottomPt = bottom[0]!.border.from.yPt;
    const continuous = (
      items: typeof group,
      startPt: number,
      endPt: number,
      coordinate: (border: ResolvedBorderSegment) => readonly [number, number],
    ): boolean => {
      const intervals = items.map(({ border }) => coordinate(border))
        .map(([start, end]) => [Math.min(start, end), Math.max(start, end)] as const)
        .sort((a, b) => a[0] - b[0]);
      if (intervals[0]?.[0] !== startPt) return false;
      let cursor = startPt;
      for (const interval of intervals) {
        if (interval[0] > cursor) return false;
        cursor = Math.max(cursor, interval[1]);
      }
      return cursor === endPt;
    };
    const isRectangle = top.every(({ border }) =>
      border.from.yPt === topPt && border.to.yPt === topPt)
      && bottom.every(({ border }) =>
        border.from.yPt === bottomPt && border.to.yPt === bottomPt)
      && left.every(({ border }) =>
        border.from.xPt === leftPt && border.to.xPt === leftPt)
      && right.every(({ border }) =>
        border.from.xPt === rightPt && border.to.xPt === rightPt)
      && continuous(top, leftPt, rightPt, (border) => [border.from.xPt, border.to.xPt])
      && continuous(bottom, leftPt, rightPt, (border) => [border.from.xPt, border.to.xPt])
      && continuous(left, topPt, bottomPt, (border) => [border.from.yPt, border.to.yPt])
      && continuous(right, topPt, bottomPt, (border) => [border.from.yPt, border.to.yPt]);
    if (!isRectangle) continue;
    const representative = group[0]!.border;
    frames.push({
      bounds: {
        xPt: leftPt,
        yPt: topPt,
        widthPt: rightPt - leftPt,
        heightPt: bottomPt - topPt,
      },
      border: {
        authoredStyle: representative.authoredStyle,
        color: representative.color,
        widthPt: representative.widthPt,
        style: representative.style,
      },
      segmentIndexes: group.map(({ index }) => index),
    });
  }
  return frames;
}

function intersectRects(left: LayoutRect, right: LayoutRect): LayoutRect | null {
  const xPt = Math.max(left.xPt, right.xPt);
  const yPt = Math.max(left.yPt, right.yPt);
  const rightPt = Math.min(left.xPt + left.widthPt, right.xPt + right.widthPt);
  const bottomPt = Math.min(left.yPt + left.heightPt, right.yPt + right.heightPt);
  return rightPt > xPt && bottomPt > yPt
    ? { xPt, yPt, widthPt: rightPt - xPt, heightPt: bottomPt - yPt }
    : null;
}

function placedChildInkBounds(
  block: TableCellBlockLayout,
  cellContentXPt: number,
  cellTopPt: number,
): LayoutRect {
  const child = block.layout;
  const targetXPt = cellContentXPt + (child.kind === 'table' ? child.flowBounds.xPt : 0);
  const targetYPt = cellTopPt + block.offsetPt + (child.kind === 'table' ? child.flowBounds.yPt : 0);
  const dxPt = targetXPt - child.flowBounds.xPt;
  const dyPt = targetYPt - child.flowBounds.yPt;
  return {
    xPt: child.inkBounds.xPt + dxPt,
    yPt: child.inkBounds.yPt + dyPt,
    widthPt: child.inkBounds.widthPt,
    heightPt: child.inkBounds.heightPt,
  };
}

/**
 * Place a rotated cell (ECMA-376 §17.4.72). The blocks keep the horizontal
 * geometry they were acquired with in a local frame whose width is the
 * physical content height and whose block axis is the physical content width;
 * §17.4.83 vAlign therefore positions the stacked lines across the cell, as
 * the rotated text flow requires. Paint and hit testing apply `transform`.
 */
function rotatedCellLayout(
  cell: TableCellLayoutInput,
  flow: CellFlowGeometry,
  cellFlowBounds: LayoutRect,
  physicalContentHeightPt: number,
  input: TableLayoutInput,
  exactOwnedSpan: boolean,
  placement: FlowBlockPlacement,
  exactClipHeightPt = cellFlowBounds.heightPt,
): TableCellLayout {
  const verticalText = cell.verticalText!;
  const physicalContent = {
    xPt: cellFlowBounds.xPt + cell.margins.leftPt,
    yPt: cellFlowBounds.yPt + cell.margins.topPt,
    widthPt: Math.max(0, cellFlowBounds.widthPt - cell.margins.leftPt - cell.margins.rightPt),
    heightPt: physicalContentHeightPt,
  };
  const localBlockExtentPt = physicalContent.widthPt;
  const localOffsetPt = flow.inkHeightPt >= localBlockExtentPt
    ? -Math.min(0, flow.inkTopPt)
    : cell.vAlign === 'center'
      ? (localBlockExtentPt - flow.inkHeightPt) / 2 - flow.inkTopPt
      : cell.vAlign === 'bottom'
        ? localBlockExtentPt - flow.inkHeightPt - flow.inkTopPt
        : -Math.min(0, flow.inkTopPt);
  const transform = verticalCellTransform(verticalText.mode, physicalContent);
  // Rotated lines can exceed the physical cell on either axis; clip them to
  // the cell like an exact row so they never paint over neighbours.
  const clipBounds = exactOwnedSpan
    ? wordExactRowVerticalClipBounds({ ...cellFlowBounds, heightPt: exactClipHeightPt }, placement.availableBounds)
    : cellFlowBounds;
  return {
    kind: 'table-cell',
    id: cell.id,
    source: cell.source,
    flowDomainId: input.flowDomainId,
    ordinaryFlow: input.ordinaryFlow,
    flowBounds: cellFlowBounds,
    inkBounds: cellFlowBounds,
    clipBounds,
    contentBounds: {
      xPt: 0,
      yPt: localOffsetPt,
      widthPt: verticalText.lineLengthPt,
      heightPt: localBlockExtentPt,
    },
    advancePt: cellFlowBounds.heightPt,
    verticalMerge: cell.verticalMerge,
    vAlign: cell.vAlign,
    ...(cell.background ? { background: cell.background } : {}),
    blocks: flow.blocks.map((block) => ({ ...block, offsetPt: localOffsetPt + block.offsetPt })),
    verticalText: { mode: verticalText.mode, transform },
  };
}

export function layoutTable(
  rawInput: TableLayoutInput,
  placement: FlowBlockPlacement,
  _services: LayoutServices,
): BlockLayoutResult<TableLayout> {
  const input = snapshotPlainData(rawInput, 'TableLayoutInput') as TableLayoutInput;
  if (input.columnWidthsPt.some((width) => !Number.isFinite(width) || width < 0)) {
    throw new TypeError('TableLayoutInput.columnWidthsPt must contain finite non-negative widths');
  }
  const flows = new Map<string, CellFlowGeometry>();
  input.rows.forEach((row) => row.cells.forEach((cell) => {
    flows.set(cell.id, resolveCellFlow(cell.verticalMerge === 'continue' ? [] : cell.blocks));
  }));
  const boundaries = resolvedBoundaries(input);
  const ruleInsetsPt = rowRuleInsetsPt(input, boundaries);
  const resolvedRows = resolveRowHeights(
    input,
    flows,
    rowBoundaryFootprintsPt(input, boundaries, ruleInsetsPt),
  );
  // ECMA-376 §17.4.80 owns the authored row height; an exact track already
  // contains its top rule, and only the table's bottom rule extends it.
  const rowHeightsPt = resolvedRows.heights;
  const widthPt = input.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const heightPt = rowHeightsPt.reduce((sum, height) => sum + height, 0);
  const yPt = placement.cursor.yPt;
  const rowXPt = input.rows.map((row) => alignedTableOriginX(
    row.alignment ?? input.alignment,
    Number.isFinite(row.indentPt) ? row.indentPt : input.indentPt,
    input.bidiVisual,
    placement,
    widthPt,
  ));
  const xPt = rowXPt[0] ?? alignedTableOriginX(
    input.alignment,
    input.indentPt,
    input.bidiVisual,
    placement,
    widthPt,
  );
  const borders = materializeBorders(input, rowXPt, yPt, rowHeightsPt, boundaries);
  const retainedCompoundFrames = compoundBorderFrames(borders);

  const rowOffsets = [0];
  for (const height of rowHeightsPt) rowOffsets.push((rowOffsets.at(-1) ?? 0) + height);
  const frame: TableRowFrame = {
    yPt,
    rowOffsetsPt: rowOffsets,
    rowHeightsPt,
    contentHeightsPt: resolvedRows.contentHeights,
    rowXPt,
    xPt,
    widthPt,
    columnOffsetsPt: columnOffsetsOf(input),
    flows,
    ruleInsetsPt,
  };
  const rows: TableRowLayout[] = input.rows.map((_row, rowIndex) => (
    materializeTableRow(input, rowIndex, frame, placement)
  ));
  const flowLeftPt = rowXPt.length > 0 ? Math.min(...rowXPt) : xPt;
  const flowRightPt = rowXPt.length > 0
    ? Math.max(...rowXPt.map((rowX) => rowX + widthPt))
    : xPt + widthPt;
  const flowBounds = {
    xPt: flowLeftPt,
    yPt,
    widthPt: Math.max(0, flowRightPt - flowLeftPt),
    heightPt,
  };
  const childInkBounds = unionLayoutRects([flowBounds, ...rows.map((row) => row.inkBounds)])
    ?? flowBounds;
  const layout: TableLayout = {
    kind: 'table',
    id: input.id,
    source: input.source,
    flowDomainId: input.flowDomainId,
    ordinaryFlow: input.ordinaryFlow,
    flowBounds,
    inkBounds: unionInkBounds(childInkBounds, borders),
    advancePt: heightPt,
    columnWidthsPt: input.columnWidthsPt,
    rows,
    borders,
    ...(retainedCompoundFrames.length
      ? { compoundBorderFrames: retainedCompoundFrames }
      : {}),
  };
  // The table layout is its own verified root, so consumers that retain only
  // `.layout` (body occurrence projection, table acquisition) hold a branded
  // immutable graph and can share it instead of copying it again.
  return snapshotPlainData({
    layout: snapshotPlainData(layout, 'TableLayoutResult.layout'),
    nextCursor: { xPt: placement.cursor.xPt, yPt: placement.cursor.yPt + heightPt },
  }, 'TableLayoutResult') as BlockLayoutResult<TableLayout>;
}

/** The row geometry layoutTable resolves for `input.rows` before it
 * materializes any row: the table top, each row's offset below it (offset
 * `r + 1` is offset `r` plus row r's height), heights, content heights and
 * origin x, the grid width and the cells' block flows by cell id. */
interface TableRowFrame {
  readonly yPt: number;
  readonly rowOffsetsPt: readonly number[];
  readonly rowHeightsPt: readonly number[];
  readonly contentHeightsPt: readonly number[];
  readonly rowXPt: readonly number[];
  readonly xPt: number;
  readonly widthPt: number;
  readonly columnOffsetsPt: readonly number[];
  readonly flows: ReadonlyMap<string, CellFlowGeometry>;
  /** Rule insets by row of the frame (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT). */
  readonly ruleInsetsPt: RowRuleInsets;
}

function columnOffsetsOf(input: TableLayoutInput): readonly number[] {
  const columnOffsets = [0];
  for (const width of input.columnWidthsPt) columnOffsets.push((columnOffsets.at(-1) ?? 0) + width);
  return columnOffsets;
}

/** Row `rowIndex` of `input` laid out in `frame`. It reads the frame at its
 * own index and, for a merge it owns, at the merge's last row. `row` is the
 * row materialized there: input row `rowIndex` itself, or a row of the same
 * structure (cells, merges, spacing; {@link laidOutTableTracks}). The other
 * rows of `input` are read only for that structure. */
function materializeTableRow(
  input: TableLayoutInput,
  rowIndex: number,
  frame: TableRowFrame,
  placement: FlowBlockPlacement,
  row: TableRowLayoutInput = input.rows[rowIndex]!,
): TableRowLayout {
  const {
    yPt, rowOffsetsPt: rowOffsets, rowHeightsPt, rowXPt, xPt, widthPt, columnOffsetsPt: columnOffsets, flows,
    ruleInsetsPt,
  } = frame;
  const columnX = (column: number) => (rowXPt[rowIndex] ?? xPt) + (input.bidiVisual
    ? widthPt - (columnOffsets[column] ?? 0)
    : (columnOffsets[column] ?? 0));
  const rowTopPt = yPt + (rowOffsets[rowIndex] ?? 0);
  const rowHeightPt = rowHeightsPt[rowIndex] ?? 0;
  const rowOriginXPt = rowXPt[rowIndex] ?? xPt;
  const rowSpacing = rowSpacingInsets(input.rows, rowIndex);
  const horizontalSpacingPt = effectiveCellSpacingPt(row);
  const cells: TableCellLayout[] = row.cells.map((cell) => {
    const opensRegion = gridMergeRole(input, row, rowIndex, cell) === 'restart';
    const lastRowIndex = opensRegion
      ? mergeEndRow(input.rows, rowIndex, cell.columnStart, cell.columnSpan)
      : rowIndex;
    const lastRowSpacing = rowSpacingInsets(input.rows, lastRowIndex);
    const cellBottomPt = yPt
      + (rowOffsets[lastRowIndex + 1] ?? rowOffsets[rowIndex + 1] ?? 0)
      - lastRowSpacing.bottomPt;
    const logicalStartX = columnX(cell.columnStart);
    const logicalEndX = columnX(
      Math.min(input.columnWidthsPt.length, cell.columnStart + cell.columnSpan),
    );
    const gridLeftPt = Math.min(logicalStartX, logicalEndX);
    const gridRightPt = Math.max(logicalStartX, logicalEndX);
    const { startPt: startInsetPt, endPt: endInsetPt } = tableCellHorizontalSpacingInsets(
      horizontalSpacingPt,
      cell.columnStart,
      cell.columnSpan,
      input.columnWidthsPt.length,
      row.spacingGeometry !== undefined,
    );
    const cellXPt = gridLeftPt + (input.bidiVisual ? endInsetPt : startInsetPt);
    const cellRightPt = gridRightPt - (input.bidiVisual ? startInsetPt : endInsetPt);
    const cellWidthPt = Math.max(0, cellRightPt - cellXPt);
    const cellTopPt = rowTopPt + rowSpacing.topPt
      + (row.spacingGeometry ? ruleInsetsPt.top[rowIndex] ?? 0 : 0);
    const cellHeightPt = opensRegion
      ? Math.max(0, cellBottomPt - cellTopPt)
      : Math.max(0, rowHeightPt - rowSpacing.topPt - rowSpacing.bottomPt
        - (row.spacingGeometry ? (ruleInsetsPt.top[rowIndex] ?? 0) + (ruleInsetsPt.bottom[rowIndex] ?? 0) : 0));
    const flow = flows.get(cell.id) ?? resolveCellFlow([]);
    // WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT: content starts below the
    // row's top rule and ends above the table's bottom rule.
    const margins = {
      ...cell.margins,
      topPt: cell.margins.topPt + (row.spacingGeometry?.borderPt ?? ruleInsetsPt.top[rowIndex] ?? 0),
      bottomPt: cell.margins.bottomPt + (row.spacingGeometry?.borderPt ?? ruleInsetsPt.bottom[lastRowIndex] ?? 0),
    };
    const physicalContentHeightPt = Math.max(
      0,
      cellHeightPt - margins.topPt - margins.bottomPt,
    );
    const exactOwnedSpan = row.heightRule === 'exact' && input.rows
      .slice(rowIndex + 1, lastRowIndex + 1)
      .every((ownedRow) => ownedRow.heightRule === 'exact');
    if (cell.verticalText && cell.verticalMerge !== 'continue') {
      return rotatedCellLayout(
        { ...cell, margins },
        flow,
        { xPt: cellXPt, yPt: cellTopPt, widthPt: cellWidthPt, heightPt: cellHeightPt },
        physicalContentHeightPt,
        input,
        exactOwnedSpan,
        placement,
        Math.max(0, cellHeightPt - (ruleInsetsPt.bottom[lastRowIndex] ?? 0)),
      );
    }
    const availableContentHeightPt = physicalContentHeightPt;
    // `word-cell-vertical-alignment-ink-block`: align the ink block, not the line boxes.
    const topInkOffsetPt = margins.topPt - Math.min(0, flow.inkTopPt);
    const inkOffsetPt = flow.inkHeightPt >= availableContentHeightPt
      ? topInkOffsetPt
      : cell.vAlign === 'center'
        ? margins.topPt + (availableContentHeightPt - flow.inkHeightPt) / 2 - flow.inkTopPt
        : cell.vAlign === 'bottom'
          ? cellHeightPt - margins.bottomPt - flow.inkHeightPt - flow.inkTopPt
          : topInkOffsetPt;
    const contentBounds = {
      xPt: cellXPt + cell.margins.leftPt + (row.spacingGeometry?.borderPt ?? 0),
      yPt: cellTopPt + inkOffsetPt,
      // Use the same grouped horizontal margins as acquisition so retained
      // geometry and line layout agree at an intrinsic equality.
      widthPt: Math.max(0, cellWidthPt - (cell.margins.leftPt + cell.margins.rightPt) - 2 * (row.spacingGeometry?.borderPt ?? 0)),
      heightPt: availableContentHeightPt,
    };
    const cellFlowBounds = { xPt: cellXPt, yPt: cellTopPt, widthPt: cellWidthPt, heightPt: cellHeightPt };
    // `word-exact-row-vertical-clip-only`
    // The authored exact box ends above the table's appended bottom rule.
    const clipBounds = cell.verticalMerge !== 'continue' && exactOwnedSpan
      ? wordExactRowVerticalClipBounds(
          { ...cellFlowBounds, heightPt: Math.max(0, cellHeightPt - (ruleInsetsPt.bottom[lastRowIndex] ?? 0)) },
          placement.availableBounds,
        )
      : undefined;
    const blocks = cell.verticalMerge === 'continue'
      ? []
      : flow.blocks.map((block) => ({
          ...block,
          offsetPt: inkOffsetPt + block.offsetPt,
        }));
    const childInk = blocks
      .map((block) => placedChildInkBounds(block, contentBounds.xPt, cellFlowBounds.yPt))
      .map((bounds) => clipBounds ? intersectRects(bounds, clipBounds) : bounds)
      .filter((bounds): bounds is LayoutRect => bounds !== null);
    const cellInkBounds = unionLayoutRects([cellFlowBounds, ...childInk]) ?? cellFlowBounds;
    return {
      kind: 'table-cell',
      id: cell.id,
      source: cell.source,
      flowDomainId: input.flowDomainId,
      ordinaryFlow: input.ordinaryFlow,
      flowBounds: cellFlowBounds,
      inkBounds: cellInkBounds,
      ...(clipBounds ? { clipBounds } : {}),
      contentBounds,
      advancePt: cellHeightPt,
      verticalMerge: cell.verticalMerge,
      vAlign: cell.vAlign,
      ...(cell.background ? { background: cell.background } : {}),
      // A projected empty owner (gridMergeRole) paints its own region: the
      // authored §17.4.32 shading fills the cell whether or not it holds
      // text. Its authored w:vMerge and suppressed (empty) blocks are kept.
      ...(cell.verticalMerge === 'continue' && opensRegion
        ? { visualMergeOwnership: 'continuation' as const }
        : {}),
      blocks,
    };
  });
  const rowBounds = { xPt: rowOriginXPt, yPt: rowTopPt, widthPt, heightPt: rowHeightPt };
  const rowInkBounds = unionLayoutRects([rowBounds, ...cells.map((cell) => cell.inkBounds)])
    ?? rowBounds;
  return {
    kind: 'table-row',
    id: row.id,
    source: row.source,
    flowDomainId: input.flowDomainId,
    ordinaryFlow: input.ordinaryFlow,
    flowBounds: rowBounds,
    inkBounds: rowInkBounds,
    advancePt: rowHeightPt,
    heightPt: rowHeightPt,
    contentHeightPt: frame.contentHeightsPt[rowIndex] ?? 0,
    ...(row.repeatedHeader ? { repeatedHeader: true } : {}),
    cells,
  };
}

/**
 * The final tracks of a laid-out table, given again to rows prepared for it.
 *
 * `row(rowIndex, candidate)` is `candidate` materialized as row `rowIndex` of
 * `laidOut`, the layout `layoutTable(input, placement)` returned: the row's
 * top, its height, its content height, its origin x, and, for a merge the
 * candidate owns, the merge's last row, are all `laidOut`'s; only the cells'
 * block flows are the candidate's. A candidate is a row whose content was
 * prepared against that layout (a cell paragraph wrapped around a positioned
 * child, a nested table placed through its page origin), with the row's
 * structure (cells, grid columns, merges, height rule, spacing). Laying it
 * out after the rows above it alone, or by itself, gives a merge continuing
 * below it the wrong end row, so its deficit, and with it every centered or
 * bottom vAlign offset, would not be the ones `laidOut` paints. Whether the
 * candidate's own content keeps those tracks is the caller's fixed point
 * (table-pagination.ts): the tracks are read from one layout and the rows
 * prepared in them are laid out again.
 *
 * The offsets are summed from `laidOut`'s row heights in row order, as
 * layoutTable offsets them, so they are its own values. Cost: the frame is
 * built once, proportional to the rows; each `row` call is proportional to
 * the candidate's cells and, for a merge it owns, that merge's span (as in
 * layoutTable).
 */
export function laidOutTableTracks(
  input: TableLayoutInput,
  laidOut: TableLayout,
  placement: FlowBlockPlacement,
): Readonly<{ row(rowIndex: number, candidate: TableRowLayoutInput): TableRowLayout }> {
  if (laidOut.rows.length !== input.rows.length) {
    throw new TypeError('laidOutTableTracks: the layout is not the input’s');
  }
  const rowHeightsPt = laidOut.rows.map((row) => row.heightPt);
  const rowOffsetsPt = [0];
  for (const height of rowHeightsPt) rowOffsetsPt.push((rowOffsetsPt.at(-1) ?? 0) + height);
  const rowXPt = laidOut.rows.map((row) => row.flowBounds.xPt);
  const frame: Omit<TableRowFrame, 'flows'> = {
    yPt: laidOut.flowBounds.yPt,
    rowOffsetsPt,
    rowHeightsPt,
    contentHeightsPt: laidOut.rows.map((row) => row.contentHeightPt),
    rowXPt,
    xPt: rowXPt[0] ?? laidOut.flowBounds.xPt,
    widthPt: input.columnWidthsPt.reduce((sum, width) => sum + width, 0),
    columnOffsetsPt: columnOffsetsOf(input),
    ruleInsetsPt: rowRuleInsetsPt(input, resolvedBoundaries(input)),
  };
  return Object.freeze({
    row(rowIndex: number, candidate: TableRowLayoutInput): TableRowLayout {
      const structural = input.rows[rowIndex];
      if (!structural || structural.cells.length !== candidate.cells.length) {
        throw new TypeError(`laidOutTableTracks: row ${rowIndex} is not the candidate’s`);
      }
      const flows = new Map(candidate.cells.map((cell) => [
        cell.id,
        resolveCellFlow(cell.verticalMerge === 'continue' ? [] : cell.blocks),
      ] as const));
      return materializeTableRow(input, rowIndex, { ...frame, flows }, placement, candidate);
    },
  });
}

/**
 * Exact row tracks of every prefix of a growing row list, at bounded cost.
 *
 * `row(rows, candidate, placement)` is the last row of
 * `layoutTable({ ...input, rows: [...rows, candidate] }, placement)`, for a
 * `rows` list that only grows by appending (a fragment's selected rows).
 * Pagination places the content of every row through such a probe
 * (table-pagination.ts), so laying the prefix out again for each row would
 * cost the square of the rows.
 *
 * Invariant: in any prefix, a row's track is settled once the row after it is
 * known, except for the deficit of a merge that continues below it:
 * - its floor, unmerged cells and spacing insets read only it and its two
 *   neighbours;
 * - a merge reads and grows only rows inside its own interval, and the
 *   merges are applied by end row (resolveRowTrack, shared with layoutTable);
 * - its collapsed-rule footprint reads only its two horizontal boundaries,
 *   and a boundary reads only the owners on its two sides: one ending above
 *   it (its terminal cell, and whether that is the table's last row) and one
 *   starting below it.
 * So the rows of `rows` before its last one are resolved once each, in order,
 * as the list grows (committed). A probe resolves only the last row (whose
 * next row is the candidate) and the candidate (the prefix's last row) on top
 * of them, without changing the committed state. Every merge continuing into
 * the candidate ends there in the prefix: its deficit lands on the candidate,
 * or, for an `exact` candidate, on the last non-exact row inside the merge,
 * which may be committed (see Arithmetic). A boundary is resolved by
 * layoutTable's own boundary resolution over the rows on its two sides, with
 * each continuation cell of the upper row that an owner covers made the owner
 * (a restart of that same cell), so the owner's terminal cell, edges and row
 * exceptions are the prefix's.
 *
 * Arithmetic: committed row tops are running sums from the table top, as
 * layoutTable offsets its rows, and a committed merge sums its interval from
 * its start, as layoutTable does, so committed rows carry layoutTable's own
 * values. A probe never re-sums committed rows. It reads a committed
 * interval (a merge still open at the candidate) and the candidate's top as
 * differences of the committed running sums, plus the growth its own
 * deficits gave committed rows (at most one per merge it closes, kept apart
 * from the committed heights) and its two own rows. A merge starting at row 0
 * that no probe deficit reaches, and a top no probe deficit lies above, are
 * then layoutTable's values; anything else equals layoutTable's up to the
 * association of the floating-point sum.
 *
 * Cost and retention: each committed row is resolved once (its cells, the
 * merges open across it and a boundary resolution over at most three rows),
 * and each committed merge sums its span once, as layoutTable does for the
 * whole table. A committed deficit landing at row t of its merge invalidates
 * the running sums from t, so they are recomputed at most over that merge's
 * span again; the recomputation is therefore bounded by the spans the
 * merges already sum. A probe resolves two rows and reads each interval and
 * its top in time proportional to the merges closing at the candidate, not to
 * the rows, whether its deficits land on the candidate, on its neighbour or
 * far above (an `exact` candidate below an `auto` merge start). One height,
 * footprint and two running sums per committed row, and the merges open into
 * the next row, are retained with the tracks, and released with them (one row
 * list's probes). A grid in which two owners claim one column of a row (a
 * malformed merge) breaks the boundary locality: a committed row's
 * collapsed-rule footprint can then depend on the terminal cell of an owner
 * that continues below the rows it reads. `row` returns null for such a list,
 * and the caller lays out the whole prefix instead (table-pagination.ts
 * wholePrefixProbe, charged to the session's acquisition budget).
 */
export interface TablePrefixTracks {
  row(
    rows: readonly TableRowLayoutInput[],
    candidate: TableRowLayoutInput,
    placement: FlowBlockPlacement,
  ): TableRowLayout | null;
}

interface CommittedRowTracks {
  /** The committed rows and the row after the last of them. */
  readonly rows: TableRowLayoutInput[];
  readonly heights: number[];
  readonly footprints: number[];
  readonly lastNonExact: number[];
  /** Σ heights[0..i−1], valid for i ≤ heightSumsValid. */
  readonly heightSums: number[];
  heightSumsValid: number;
  /** Σ (heights + footprints)[0..i−1], valid for i ≤ offsetsValid. */
  readonly offsets: number[];
  offsetsValid: number;
  open: readonly OpenMergeTrack[];
  /** Continuation cells of the last committed row that an owner covers. */
  covered: ReadonlySet<TableCellLayoutInput>;
  regular: boolean;
}

/** The continuation cells of `row` the merges open into it cover, and
 * whether no column of the row is claimed by two owners. `above` is the row
 * above it, so a projected empty owner claims its columns (gridMergeRole). */
function rowCoverage(
  row: TableRowLayoutInput,
  above: TableRowLayoutInput | undefined,
  open: readonly OpenMergeTrack[],
  columnCount: number,
  segmentOpeningLogicalRowIndex: TableLayoutInput['segmentOpeningLogicalRowIndex'],
): Readonly<{ covered: ReadonlySet<TableCellLayoutInput>; regular: boolean }> {
  const claimed = new Set<number>();
  let regular = true;
  const claim = (cell: TableCellLayoutInput) => {
    const end = Math.min(columnCount, cell.columnStart + cell.columnSpan);
    for (let column = Math.max(0, cell.columnStart); column < end; column += 1) {
      if (claimed.has(column)) regular = false;
      claimed.add(column);
    }
  };
  const covered = new Set<TableCellLayoutInput>();
  for (const merge of open) {
    claim(merge.cell);
    const continuation = row.cells.find((cell) => continuesMerge(cell, merge.cell));
    if (continuation) covered.add(continuation);
  }
  for (const cell of row.cells) {
    if (projectedMergeRole(segmentOpeningLogicalRowIndex, above, row, cell) !== 'continue') claim(cell);
  }
  return { covered, regular };
}

export function tablePrefixTracks(input: TableLayoutInput): TablePrefixTracks {
  const columnCount = input.columnWidthsPt.length;
  const flows = new WeakMap<TableCellLayoutInput, CellFlowGeometry>();
  const flowOf = (cell: TableCellLayoutInput): CellFlowGeometry => {
    let flow = flows.get(cell);
    if (!flow) {
      flow = resolveCellFlow(cell.verticalMerge === 'continue' ? [] : cell.blocks);
      flows.set(cell, flow);
    }
    return flow;
  };
  const emptyTracks = (): CommittedRowTracks => ({
    rows: [],
    heights: [],
    footprints: [],
    lastNonExact: [],
    heightSums: [0],
    heightSumsValid: 0,
    offsets: [0],
    offsetsValid: 0,
    open: [],
    covered: new Set<TableCellLayoutInput>(),
    regular: true,
  });
  let tracks = emptyTracks();
  const ensureHeightSums = (upTo: number) => {
    for (let index = tracks.heightSumsValid; index < upTo; index += 1) {
      tracks.heightSums[index + 1] = tracks.heightSums[index]! + tracks.heights[index]!;
    }
    tracks.heightSumsValid = Math.max(tracks.heightSumsValid, upTo);
  };
  const ensureOffsets = (upTo: number) => {
    for (let index = tracks.offsetsValid; index < upTo; index += 1) {
      tracks.offsets[index + 1] = tracks.offsets[index]! + (tracks.heights[index]! + tracks.footprints[index]!);
    }
    tracks.offsetsValid = Math.max(tracks.offsetsValid, upTo);
  };
  const committedStore: RowTrackStore = {
    get length() { return tracks.heights.length; },
    push(heightPt, _contentHeightPt, exact) {
      tracks.lastNonExact.push(exact ? (tracks.lastNonExact.at(-1) ?? -1) : tracks.heights.length);
      tracks.heights.push(heightPt);
    },
    sum(start, end) {
      // Each merge ends once, so summing its interval as layoutTable does
      // (from its start) costs its span once.
      if (start > 0) {
        let sumPt = 0;
        for (let rowIndex = start; rowIndex <= end; rowIndex += 1) sumPt += tracks.heights[rowIndex]!;
        return sumPt;
      }
      ensureHeightSums(end + 1);
      return tracks.heightSums[end + 1]!;
    },
    grow(rowIndex, deficitPt) {
      tracks.heights[rowIndex] = tracks.heights[rowIndex]! + deficitPt;
      tracks.heightSumsValid = Math.min(tracks.heightSumsValid, rowIndex);
      tracks.offsetsValid = Math.min(tracks.offsetsValid, rowIndex);
    },
    lastNonExact: (rowIndex) => tracks.lastNonExact[rowIndex] ?? -1,
    raiseContentHeight() {},
  };
  // The collapsed-rule footprint of `row`, between `above` and `next`.
  // The rows around `row` with the merges it continues restored to their
  // owners, so a merge's terminal rule takes part in `row`'s boundaries.
  const neighbourhood = (
    above: Readonly<{ row: TableRowLayoutInput; covered: ReadonlySet<TableCellLayoutInput> }> | null,
    row: TableRowLayoutInput,
    next: TableRowLayoutInput | undefined,
  ): Readonly<{ rows: TableRowLayoutInput[]; index: number }> => {
    const owners = above && above.covered.size > 0
      ? {
          ...above.row,
          cells: above.row.cells.map((cell) => (
            above.covered.has(cell) ? { ...cell, verticalMerge: 'restart' as const } : cell
          )),
        }
      : above?.row;
    return { rows: [...(owners ? [owners] : []), row, ...(next ? [next] : [])], index: owners ? 1 : 0 };
  };
  const footprintOf = (
    above: Readonly<{ row: TableRowLayoutInput; covered: ReadonlySet<TableCellLayoutInput> }> | null,
    row: TableRowLayoutInput,
    next: TableRowLayoutInput | undefined,
  ): number => {
    const { rows, index } = neighbourhood(above, row, next);
    return tableRowBoundaryFootprintsPt({ ...input, rows })[index] ?? 0;
  };
  const commitNext = (rows: readonly TableRowLayoutInput[]) => {
    const rowIndex = tracks.heights.length;
    const row = rows[rowIndex]!;
    const next = rows[rowIndex + 1]!;
    const above = rowIndex === 0 ? null : { row: rows[rowIndex - 1]!, covered: tracks.covered };
    const coverage = rowCoverage(row, above?.row, tracks.open, columnCount, input.segmentOpeningLogicalRowIndex);
    tracks.open = resolveRowTrack(
      committedStore, above ? [above.row, row, next] : [row, next], above ? 1 : 0, tracks.open, flowOf,
      input.segmentOpeningLogicalRowIndex,
    );
    tracks.footprints.push(footprintOf(above, row, next));
    tracks.covered = coverage.covered;
    tracks.regular &&= coverage.regular;
    tracks.rows[rowIndex] = row;
    tracks.rows[rowIndex + 1] = next;
  };
  // Commit every row of `rows` but its last; a list that is not the committed
  // one grown by appending starts over.
  const sync = (rows: readonly TableRowLayoutInput[]) => {
    const target = Math.max(0, rows.length - 1);
    const committed = tracks.heights.length;
    if (committed > target || (committed > 0 && (
      rows[committed - 1] !== tracks.rows[committed - 1] || rows[committed] !== tracks.rows[committed]
    ))) tracks = emptyTracks();
    while (tracks.heights.length < target) commitNext(rows);
  };
  const resolve = (rows: readonly TableRowLayoutInput[], candidate: TableRowLayoutInput) => {
    sync(rows);
    const count = rows.length;
    const base = tracks.heights.length;
    const heights: number[] = [];
    const lastNonExact: number[] = [];
    let contentHeightPt = 0;
    // Committed rows a deficit of this probe grew, and their growth: at most
    // one entry per merge the probe closes, never written to the committed
    // state.
    const landed = new Map<number, number>();
    // A row of the probe's own (rowIndex ≥ base).
    const heightAt = (rowIndex: number) => heights[rowIndex - base]!;
    // The growth landed on committed rows start..end−1.
    const landedIn = (start: number, end: number) => {
      let growthPt = 0;
      for (const [rowIndex, deficitPt] of landed) {
        if (rowIndex >= start && rowIndex < end) growthPt += deficitPt;
      }
      return growthPt;
    };
    const lastNonExactAt = (rowIndex: number) => (
      rowIndex >= base ? lastNonExact[rowIndex - base]! : tracks.lastNonExact[rowIndex] ?? -1
    );
    const store: RowTrackStore = {
      get length() { return base + heights.length; },
      push(heightPt, rowContentHeightPt, exact) {
        const rowIndex = base + heights.length;
        lastNonExact.push(exact ? (rowIndex === 0 ? -1 : lastNonExactAt(rowIndex - 1)) : rowIndex);
        heights.push(heightPt);
        if (rowIndex === count) contentHeightPt = rowContentHeightPt;
      },
      sum(start, end) {
        // Committed rows as a difference of running sums plus the growth
        // landed on them, then the probe's own rows (at most two): no
        // committed row is summed again.
        let sumPt = 0;
        const committedEnd = Math.min(end + 1, base);
        if (start < committedEnd) {
          ensureHeightSums(committedEnd);
          sumPt = tracks.heightSums[committedEnd]! - tracks.heightSums[start]!;
          if (landed.size > 0) sumPt += landedIn(start, committedEnd);
        }
        for (let rowIndex = Math.max(start, base); rowIndex <= end; rowIndex += 1) sumPt += heightAt(rowIndex);
        return sumPt;
      },
      grow(rowIndex, deficitPt) {
        if (rowIndex >= base) heights[rowIndex - base] = heights[rowIndex - base]! + deficitPt;
        else landed.set(rowIndex, (landed.get(rowIndex) ?? 0) + deficitPt);
      },
      lastNonExact: lastNonExactAt,
      raiseContentHeight(rowIndex, requiredPt) {
        if (rowIndex === count) contentHeightPt = Math.max(contentHeightPt, requiredPt);
      },
    };
    let open = tracks.open;
    let above = count >= 2 ? { row: rows[count - 2]!, covered: tracks.covered } : null;
    let regular = tracks.regular;
    let lastFootprintPt = 0;
    if (count >= 1) {
      const last = rows[count - 1]!;
      const coverage = rowCoverage(last, above?.row, open, columnCount, input.segmentOpeningLogicalRowIndex);
      regular &&= coverage.regular;
      open = resolveRowTrack(
        store, above ? [above.row, last, candidate] : [last, candidate], above ? 1 : 0, open, flowOf,
        input.segmentOpeningLogicalRowIndex,
      );
      lastFootprintPt = footprintOf(above, last, candidate);
      above = { row: last, covered: coverage.covered };
    }
    regular &&= rowCoverage(candidate, above?.row, open, columnCount, input.segmentOpeningLogicalRowIndex).regular;
    resolveRowTrack(
      store, above ? [above.row, candidate] : [candidate], above ? 1 : 0, open, flowOf,
      input.segmentOpeningLogicalRowIndex,
    );
    if (!regular) return null;
    const candidateFootprintPt = footprintOf(above, candidate, undefined);
    const candidateNeighbourhood = neighbourhood(above, candidate, undefined);
    const candidateInsets = tableRowRuleInsetsPt({ ...input, rows: candidateNeighbourhood.rows });
    const footprintAt = (rowIndex: number) => (
      rowIndex >= base ? lastFootprintPt : tracks.footprints[rowIndex]!
    );
    // The committed top of the probe's first own row, the growth landed above
    // it (every landed row is committed), then the probe's own rows above the
    // candidate (at most one).
    ensureOffsets(base);
    let topOffsetPt = tracks.offsets[base]!;
    if (landed.size > 0) topOffsetPt += landedIn(0, base);
    for (let rowIndex = base; rowIndex < count; rowIndex += 1) {
      topOffsetPt += heightAt(rowIndex) + footprintAt(rowIndex);
    }
    return {
      topOffsetPt,
      heightPt: heightAt(count) + candidateFootprintPt,
      contentHeightPt,
      ruleTopPt: candidateInsets.top[candidateNeighbourhood.index] ?? 0,
      ruleBottomPt: candidateInsets.bottom[candidateNeighbourhood.index] ?? 0,
    };
  };
  const prefixTracks: TablePrefixTracks = {
    row(rows, candidate, placement) {
      const resolved = resolve(rows, candidate);
      if (!resolved) return null;
      // The candidate after its upper neighbour reads, as the prefix's last
      // row, only its own track and that neighbour's spacing.
      const local = rows.length > 0 ? [rows[rows.length - 1]!, candidate] : [candidate];
      const localInput: TableLayoutInput = { ...input, rows: local };
      const widthPt = input.columnWidthsPt.reduce((sum, width) => sum + width, 0);
      const rowXPt = local.map((row) => alignedTableOriginX(
        row.alignment ?? input.alignment,
        Number.isFinite(row.indentPt) ? row.indentPt : input.indentPt,
        input.bidiVisual,
        placement,
        widthPt,
      ));
      const rowOffsetsPt = local.map(() => resolved.topOffsetPt);
      rowOffsetsPt.push(resolved.topOffsetPt + resolved.heightPt);
      return materializeTableRow(localInput, local.length - 1, {
        yPt: placement.cursor.yPt,
        rowOffsetsPt,
        rowHeightsPt: local.map(() => resolved.heightPt),
        contentHeightsPt: local.map(() => resolved.contentHeightPt),
        rowXPt,
        xPt: rowXPt[0]!,
        widthPt,
        columnOffsetsPt: columnOffsetsOf(input),
        flows: new Map(candidate.cells.map((cell) => [cell.id, flowOf(cell)] as const)),
        ruleInsetsPt: {
          top: local.map((_row, index) => index === local.length - 1 ? resolved.ruleTopPt : 0),
          bottom: local.map((_row, index) => index === local.length - 1 ? resolved.ruleBottomPt : 0),
        },
      }, placement);
    },
  };
  return Object.freeze(prefixTracks);
}
