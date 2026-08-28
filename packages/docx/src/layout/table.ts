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
 * spacing; an internal boundary is one shared gap split symmetrically between
 * its two cells. When adjacent rows differ, the larger minimum owns the shared
 * boundary so neither row's constraint is weakened.
 */
function rowSpacingInsets(
  rows: readonly TableRowLayoutInput[],
  rowIndex: number,
): RowSpacingInsets {
  const currentPt = effectiveCellSpacingPt(rows[rowIndex]);
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
  return spacing.topPt
    + input.margins.topPt
    + cellContentRequiredHeightPt(flow)
    + input.margins.bottomPt
    + spacing.bottomPt;
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

function semanticRowFloor(row: TableRowLayoutInput): number {
  if (row.heightRule === 'exact') {
    // Compatibility-owned exact-row floor.
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

function resolveRowHeights(
  rows: readonly TableRowLayoutInput[],
  flows: ReadonlyMap<string, CellFlowGeometry>,
  boundaryFootprintsPt: readonly number[],
): Readonly<{ heights: readonly number[]; contentHeights: readonly number[] }> {
  const heights = rows.map((row) => semanticRowFloor(row));
  const contentHeights = rows.map((row, rowIndex) => Math.max(
    0,
    ...row.cells
      .filter((cell) => cell.verticalMerge !== 'continue')
      .map((cell) => {
        const endRowIndex = cell.verticalMerge === 'restart'
          ? mergeEndRow(rows, rowIndex, cell.columnStart, cell.columnSpan)
          : rowIndex;
        const startInsets = rowSpacingInsets(rows, rowIndex);
        const endInsets = rowSpacingInsets(rows, endRowIndex);
        return cellRequiredHeight(
          cell,
          flows.get(cell.id) ?? resolveCellFlow([]),
          { topPt: startInsets.topPt, bottomPt: endInsets.bottomPt },
        );
      }),
  ));
  rows.forEach((row, rowIndex) => {
    const spacing = rowSpacingInsets(rows, rowIndex);
    for (const cell of row.cells) {
      if (cell.verticalMerge !== 'none') continue;
      const required = cellRequiredHeight(cell, flows.get(cell.id) ?? resolveCellFlow([]), spacing);
      if (row.heightRule !== 'exact') heights[rowIndex] = Math.max(heights[rowIndex] ?? 0, required);
    }
  });

  // A merged owner is one interval constraint over row tracks. ECMA-376 defines
  // the merged region but not deficit distribution. The terminal-growable greedy
  // policy makes the minimum total change, preserves earlier boundaries, reuses
  // prior interval growth, and never violates an exact track.
  const constraints: Array<{
    start: number;
    end: number;
    requiredPt: number;
  }> = [];
  rows.forEach((row, rowIndex) => {
    for (const cell of row.cells) {
      if (cell.verticalMerge !== 'restart') continue;
      constraints.push({
        start: rowIndex,
        end: mergeEndRow(rows, rowIndex, cell.columnStart, cell.columnSpan),
        requiredPt: cellRequiredHeight(
          cell,
          flows.get(cell.id) ?? resolveCellFlow([]),
          {
            topPt: rowSpacingInsets(rows, rowIndex).topPt,
            bottomPt: rowSpacingInsets(
              rows,
              mergeEndRow(rows, rowIndex, cell.columnStart, cell.columnSpan),
            ).bottomPt,
          },
        ),
      });
    }
  });
  constraints.sort((left, right) => left.end - right.end || left.start - right.start);
  for (const constraint of constraints) {
    let currentPt = 0;
    for (let rowIndex = constraint.start; rowIndex <= constraint.end; rowIndex += 1) {
      currentPt += heights[rowIndex] ?? 0;
    }
    const deficitPt = constraint.requiredPt - currentPt;
    if (deficitPt <= 0) continue;
    for (let rowIndex = constraint.end; rowIndex >= constraint.start; rowIndex -= 1) {
      if (rows[rowIndex]?.heightRule === 'exact') continue;
      heights[rowIndex] = (heights[rowIndex] ?? 0) + deficitPt;
      break;
    }
  }
  rows.forEach((row, rowIndex) => {
    if (row.heightRule === 'exact') return;
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
      if (cell.verticalMerge === 'continue') continue;
      const lastRowIndex = cell.verticalMerge === 'restart'
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

function rowBoundaryFootprintsPt(
  input: TableLayoutInput,
  boundaries: ReturnType<typeof resolvedBoundaries>,
): readonly number[] {
  const widthsPt = collapsedHorizontalBoundaryWidthsPt(input, boundaries);
  return input.rows.map((row, rowIndex) => (
    row.heightRule === 'exact'
      ? 0
      : wordCollapsedBorderRowTrackFootprintPt(
          widthsPt[rowIndex] ?? 0,
          widthsPt[rowIndex + 1] ?? 0,
        )
  ));
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
    edge: resolved.edge,
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
    for (const segment of merged) {
      segments.push(borderSegment(
        segment.resolved,
        { xPt: segment.leftPt, yPt: rowY(boundary) },
        { xPt: segment.rightPt, yPt: rowY(boundary) },
      ));
    }
  });
  boundaries.vertical.forEach((rows, boundary) => {
    rows.forEach((resolved, row) => {
      if (effectiveCellSpacingPt(input.rows[row]) > 0) return;
      if (!resolved) return;
      segments.push(borderSegment(
        resolved,
        { xPt: columnX(row, boundary), yPt: rowY(row) },
        { xPt: columnX(row, boundary), yPt: rowY(row + 1) },
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
      if (cell.verticalMerge === 'continue') continue;
      const lastRowIndex = cell.verticalMerge === 'restart'
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
  const resolvedRows = resolveRowHeights(
    input.rows,
    flows,
    rowBoundaryFootprintsPt(input, boundaries),
  );
  // ECMA-376 §17.4.80 owns the authored row height. Word's page-local collapsed
  // boundary footprint is added only to non-exact tracks above; exact tracks
  // already define the complete row box.
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

  const columnOffsets = [0];
  for (const width of input.columnWidthsPt) columnOffsets.push((columnOffsets.at(-1) ?? 0) + width);
  const rowOffsets = [0];
  for (const height of rowHeightsPt) rowOffsets.push((rowOffsets.at(-1) ?? 0) + height);
  const columnX = (rowIndex: number, column: number) => (rowXPt[rowIndex] ?? xPt) + (input.bidiVisual
    ? widthPt - (columnOffsets[column] ?? 0)
    : (columnOffsets[column] ?? 0));

  const rows: TableRowLayout[] = input.rows.map((row, rowIndex) => {
    const rowTopPt = yPt + (rowOffsets[rowIndex] ?? 0);
    const rowHeightPt = rowHeightsPt[rowIndex] ?? 0;
    const rowOriginXPt = rowXPt[rowIndex] ?? xPt;
    const rowSpacing = rowSpacingInsets(input.rows, rowIndex);
    const horizontalSpacingPt = effectiveCellSpacingPt(row);
    const cells: TableCellLayout[] = row.cells.map((cell) => {
      const lastRowIndex = cell.verticalMerge === 'restart'
        ? mergeEndRow(input.rows, rowIndex, cell.columnStart, cell.columnSpan)
        : rowIndex;
      const lastRowSpacing = rowSpacingInsets(input.rows, lastRowIndex);
      const cellBottomPt = yPt
        + (rowOffsets[lastRowIndex + 1] ?? rowOffsets[rowIndex + 1] ?? 0)
        - lastRowSpacing.bottomPt;
      const logicalStartX = columnX(rowIndex, cell.columnStart);
      const logicalEndX = columnX(
        rowIndex,
        Math.min(input.columnWidthsPt.length, cell.columnStart + cell.columnSpan),
      );
      const gridLeftPt = Math.min(logicalStartX, logicalEndX);
      const gridRightPt = Math.max(logicalStartX, logicalEndX);
      const { startPt: startInsetPt, endPt: endInsetPt } = tableCellHorizontalSpacingInsets(
        horizontalSpacingPt,
        cell.columnStart,
        cell.columnSpan,
        input.columnWidthsPt.length,
      );
      const cellXPt = gridLeftPt + (input.bidiVisual ? endInsetPt : startInsetPt);
      const cellRightPt = gridRightPt - (input.bidiVisual ? startInsetPt : endInsetPt);
      const cellWidthPt = Math.max(0, cellRightPt - cellXPt);
      const cellTopPt = rowTopPt + rowSpacing.topPt;
      const cellHeightPt = cell.verticalMerge === 'restart'
        ? Math.max(0, cellBottomPt - cellTopPt)
        : Math.max(0, rowHeightPt - rowSpacing.topPt - rowSpacing.bottomPt);
      const flow = flows.get(cell.id) ?? resolveCellFlow([]);
      const availableContentHeightPt = Math.max(
        0,
        cellHeightPt - cell.margins.topPt - cell.margins.bottomPt,
      );
      const topInkOffsetPt = cell.margins.topPt - Math.min(0, flow.inkTopPt);
      const inkOffsetPt = flow.inkHeightPt >= availableContentHeightPt
        ? topInkOffsetPt
        : cell.vAlign === 'center'
          ? cell.margins.topPt + (availableContentHeightPt - flow.inkHeightPt) / 2 - flow.inkTopPt
          : cell.vAlign === 'bottom'
            ? cellHeightPt - cell.margins.bottomPt - flow.inkHeightPt - flow.inkTopPt
            : topInkOffsetPt;
      const contentBounds = {
        xPt: cellXPt + cell.margins.leftPt,
        yPt: cellTopPt + inkOffsetPt,
        widthPt: Math.max(0, cellWidthPt - cell.margins.leftPt - cell.margins.rightPt),
        heightPt: availableContentHeightPt,
      };
      const cellFlowBounds = { xPt: cellXPt, yPt: cellTopPt, widthPt: cellWidthPt, heightPt: cellHeightPt };
      const exactOwnedSpan = cell.verticalMerge !== 'continue'
        && input.rows
          .slice(rowIndex, lastRowIndex + 1)
          .every((ownedRow) => ownedRow.heightRule === 'exact');
      const clipBounds = exactOwnedSpan
        ? wordExactRowVerticalClipBounds(
            cellFlowBounds,
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
      contentHeightPt: resolvedRows.contentHeights[rowIndex] ?? 0,
      ...(row.repeatedHeader ? { repeatedHeader: true } : {}),
      cells,
    };
  });
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
  return snapshotPlainData({
    layout,
    nextCursor: { xPt: placement.cursor.xPt, yPt: placement.cursor.yPt + heightPt },
  }, 'TableLayoutResult') as BlockLayoutResult<TableLayout>;
}
