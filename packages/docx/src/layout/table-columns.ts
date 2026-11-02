import type {
  TableColumnCellConstraint,
  TableColumnLayoutInput,
  TableColumnRowConstraint,
  TablePreferredWidthConstraint,
} from './types.js';
import { exactLengthKeyFromNumber, type ExactLengthKey } from './exact-length.js';

const EPSILON_PT = 1e-9;

export interface TableCellHorizontalSpacingInsets {
  readonly startPt: number;
  readonly endPt: number;
}

/** §17.4.43/.44/.45 spacing lives inside the table grid: an outer edge owns
 * the full spacing and two adjacent cells share one inner gap. This projection
 * is shared by intrinsic constraints, child acquisition, and final cell boxes
 * so each stage uses the same content width. */
export function tableCellHorizontalSpacingInsets(
  spacingPt: number,
  columnStart: number,
  columnSpan: number,
  columnCount: number,
): TableCellHorizontalSpacingInsets {
  const spacing = Number.isFinite(spacingPt) ? Math.max(0, spacingPt) : 0;
  const start = Math.max(0, columnStart);
  const end = start + Math.max(1, columnSpan);
  return {
    startPt: start === 0 ? spacing : spacing / 2,
    endPt: end >= Math.max(0, columnCount) ? spacing : spacing / 2,
  };
}

function finiteNonNegative(value: number | null | undefined): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(0, value)
    : 0;
}

function cleanWidths(widths: readonly number[]): number[] {
  return widths.map((width) => (Math.abs(width) <= EPSILON_PT ? 0 : width));
}

function requiredColumnCount(input: TableColumnLayoutInput): number {
  let count = input.gridWidthsPt.length;
  for (const row of input.rows) {
    for (const cell of row.cells) {
      count = Math.max(count, cell.columnStart + Math.max(1, cell.columnSpan));
    }
    const lastCellEnd = row.cells.reduce(
      (maximum, cell) => Math.max(maximum, cell.columnStart + Math.max(1, cell.columnSpan)),
      row.before?.columnSpan ?? 0,
    );
    count = Math.max(count, lastCellEnd + (row.after?.columnSpan ?? 0));
  }
  return count;
}

function spanSum(widths: readonly number[], start: number, span: number): number {
  let total = 0;
  const end = Math.min(widths.length, start + Math.max(1, span));
  for (let column = Math.max(0, start); column < end; column += 1) {
    total += widths[column] ?? 0;
  }
  return total;
}

function preferredPoints(
  preferred: TablePreferredWidthConstraint | null,
  percentageBasePt: number,
): number | null {
  if (!preferred) return null;
  return preferred.kind === 'pct'
    ? finiteNonNegative(preferred.value) * percentageBasePt
    : finiteNonNegative(preferred.value);
}

function setFirstRowSpan(
  widths: number[],
  start: number,
  span: number,
  preferredPt: number,
): void {
  const safeStart = Math.max(0, start);
  const safeSpan = Math.max(1, Math.min(span, widths.length - safeStart));
  if (safeSpan <= 0) return;
  const currentPt = spanSum(widths, safeStart, safeSpan);
  if (currentPt <= EPSILON_PT) {
    widths[safeStart + safeSpan - 1] = preferredPt;
    return;
  }
  // The guidance does not assign a first-row spanning delta to an individual
  // grid track. Preserve the authored tblGrid proportions inside that span;
  // subsequent-row disagreements have a separately documented last-ending
  // column rule below.
  const scale = preferredPt / currentPt;
  for (let column = safeStart; column < safeStart + safeSpan; column += 1) {
    widths[column] = (widths[column] ?? 0) * scale;
  }
}

function growLaterRowSpan(
  widths: number[],
  start: number,
  span: number,
  preferredPt: number,
): void {
  const safeStart = Math.max(0, start);
  const safeSpan = Math.max(1, Math.min(span, widths.length - safeStart));
  if (safeSpan <= 0) return;
  const currentPt = spanSum(widths, safeStart, safeSpan);
  if (preferredPt <= currentPt + EPSILON_PT) return;
  // §17.18.87's fixed guidance assigns a later-row disagreement to the last
  // grid column ending at that boundary. Keeping that exact ownership avoids a
  // sample-dependent distribution rule for merged cells.
  widths[safeStart + safeSpan - 1] += preferredPt - currentPt;
}

function rowConstraints(
  row: TableColumnRowConstraint,
  columnCount: number,
): Array<Readonly<{
  start: number;
  span: number;
  preferred: TablePreferredWidthConstraint | null;
}>> {
  const constraints: Array<Readonly<{
    start: number;
    span: number;
    preferred: TablePreferredWidthConstraint | null;
  }>> = [];
  if (row.before && row.before.columnSpan > 0) {
    constraints.push({ start: 0, span: row.before.columnSpan, preferred: row.before.preferredWidth });
  }
  for (const cell of row.cells) {
    constraints.push({
      start: cell.columnStart,
      span: cell.columnSpan,
      preferred: cell.preferredWidth,
    });
  }
  if (row.after && row.after.columnSpan > 0) {
    constraints.push({
      start: Math.max(0, columnCount - row.after.columnSpan),
      span: row.after.columnSpan,
      preferred: row.after.preferredWidth,
    });
  }
  return constraints;
}

function fixedWidths(input: TableColumnLayoutInput, columnCount: number): number[] {
  const widths = Array.from({ length: columnCount }, (_unused, column) => (
    finiteNonNegative(input.gridWidthsPt[column] ?? 0)
  ));
  const initialTotalPt = widths.reduce((sum, width) => sum + width, 0);
  const percentageBasePt = input.tablePreferredWidthPt
    ?? (initialTotalPt > 0 ? initialTotalPt : finiteNonNegative(input.availableWidthPt));

  input.rows.forEach((row, rowIndex) => {
    for (const constraint of rowConstraints(row, columnCount)) {
      const preferredPt = preferredPoints(constraint.preferred, percentageBasePt);
      if (preferredPt === null) continue;
      if (rowIndex === 0) {
        setFirstRowSpan(widths, constraint.start, constraint.span, preferredPt);
      } else {
        growLaterRowSpan(widths, constraint.start, constraint.span, preferredPt);
      }
    }
  });

  if (input.tablePreferredWidthPt === null) {
    resolveSingleColumnPercentages(widths, input.rows);
  }

  const targetPt = input.tablePreferredWidthPt;
  const totalPt = widths.reduce((sum, width) => sum + width, 0);
  if (targetPt !== null && targetPt >= 0 && totalPt <= EPSILON_PT && widths.length > 0) {
    // tblW still defines the total table width when every declared grid track
    // has zero width. The standard supplies no internal ratio in that case, so
    // equal tracks are the only source-independent deterministic allocation.
    return widths.map(() => targetPt / widths.length);
  }
  if (targetPt !== null && targetPt >= 0 && totalPt > EPSILON_PT) {
    if (input.layout === 'autofit' && targetPt < totalPt) {
      const cells = input.rows.flatMap((row) => row.cells);
      if (cells.some((cell) => cell.noWrap === true && cell.preferredWidth?.kind === 'dxa')) {
        // §17.4.29's dxa priority also applies to the initial tblW fit. Use
        // authored track geometry before proportional shrink loses it, and
        // reduce unprotected neighbours to their content minima first. Fixed
        // layout and AutoFit without dxa noWrap retain the proportional pass.
        const { minimums } = singleColumnBounds(input.rows, columnCount, totalPt);
        return fitToAvailableWidth(widths, minimums, cells, targetPt, 0);
      }
    }
    const scale = targetPt / totalPt;
    return widths.map((width) => width * scale);
  }
  return widths;
}

/**
 * tcW percentages are relative to the resulting table width (§17.4.71), so a
 * percentage cannot in general be converted from the initial tblGrid total.
 * For cells spanning one grid track, solve the minimal fixed point exactly:
 * each track is max(authored/requested points, percentage * final total).
 *
 * Spanning percentage constraints remain represented by their first fixed
 * pass because the standard does not prescribe how to distribute a span's
 * circular delta over its internal grid tracks.
 */
function resolveSingleColumnPercentages(
  widths: number[],
  rows: readonly TableColumnRowConstraint[],
): void {
  const percentages = new Array<number>(widths.length).fill(0);
  for (const row of rows) {
    for (const cell of row.cells) {
      if (cell.columnSpan !== 1 || cell.preferredWidth?.kind !== 'pct') continue;
      const column = cell.columnStart;
      if (column < 0 || column >= widths.length) continue;
      percentages[column] = Math.max(
        percentages[column] ?? 0,
        finiteNonNegative(cell.preferredWidth.value),
      );
    }
  }

  let totalPt = widths.reduce((sum, width) => sum + width, 0);
  let active = percentages.map((percentage, column) => (
    percentage * totalPt > widths[column]! + EPSILON_PT
  ));
  const visited = new Set<string>();

  while (active.some(Boolean)) {
    const signature = active.map((value) => (value ? '1' : '0')).join('');
    if (visited.has(signature)) return;
    visited.add(signature);

    let fixedPt = 0;
    let activeFraction = 0;
    for (let column = 0; column < widths.length; column += 1) {
      if (active[column]) activeFraction += percentages[column] ?? 0;
      else fixedPt += widths[column] ?? 0;
    }
    // A set of preferences at or above 100% is internally inconsistent unless
    // it already spans the whole table. Widths are preferred, so retain the
    // finite fixed-pass result instead of inventing an unbounded table size.
    if (activeFraction >= 1 - EPSILON_PT) return;

    totalPt = fixedPt / (1 - activeFraction);
    const next = percentages.map((percentage, column) => (
      percentage * totalPt > widths[column]! + EPSILON_PT
    ));
    if (next.every((value, column) => value === active[column])) {
      for (let column = 0; column < widths.length; column += 1) {
        if (active[column]) widths[column] = percentages[column]! * totalPt;
      }
      return;
    }
    active = next;
  }
}

function singleColumnBounds(
  rows: readonly TableColumnRowConstraint[],
  columnCount: number,
  percentageBasePt: number,
): Readonly<{ minimums: number[]; maximums: number[] }> {
  const minimums = new Array<number>(columnCount).fill(0);
  const maximums = new Array<number>(columnCount).fill(0);
  for (const row of rows) {
    for (const cell of row.cells) {
      if (cell.columnSpan !== 1 || cell.columnStart < 0 || cell.columnStart >= columnCount) continue;
      minimums[cell.columnStart] = Math.max(
        minimums[cell.columnStart] ?? 0,
        finiteNonNegative(cell.minContentWidthPt),
      );
      maximums[cell.columnStart] = Math.max(
        maximums[cell.columnStart] ?? 0,
        finiteNonNegative(cell.maxContentWidthPt),
      );
    }
  }

  // §17.18.87: the first preferred width in a one-column track overrides that
  // column's maximum content width. Minimum content remains a hard lower bound
  // until the final forced-line-break step.
  const preferredSeen = new Array<boolean>(columnCount).fill(false);
  for (const row of rows) {
    for (const cell of row.cells) {
      const column = cell.columnStart;
      if (
        cell.columnSpan !== 1
        || column < 0
        || column >= columnCount
        || preferredSeen[column]
        || cell.preferredWidth === null
      ) continue;
      const preferredPt = preferredPoints(cell.preferredWidth, percentageBasePt);
      if (preferredPt === null) continue;
      maximums[column] = Math.max(minimums[column] ?? 0, preferredPt);
      preferredSeen[column] = true;
    }
  }
  for (let column = 0; column < columnCount; column += 1) {
    maximums[column] = Math.max(minimums[column] ?? 0, maximums[column] ?? 0);
  }
  return { minimums, maximums };
}

function shrinkOutsideSpan(
  widths: number[],
  minimums: readonly number[],
  start: number,
  span: number,
  requestedPt: number,
): number {
  const end = Math.min(widths.length, start + span);
  const outside = widths.map((_width, column) => column).filter((column) => (
    column < start || column >= end
  ));
  const slack = outside.map((column) => Math.max(0, widths[column]! - (minimums[column] ?? 0)));
  const totalSlackPt = slack.reduce((sum, value) => sum + value, 0);
  const transferredPt = Math.min(requestedPt, totalSlackPt);
  if (transferredPt <= EPSILON_PT || totalSlackPt <= EPSILON_PT) return 0;
  outside.forEach((column, index) => {
    widths[column] -= transferredPt * ((slack[index] ?? 0) / totalSlackPt);
  });
  return transferredPt;
}

function growSpan(widths: number[], start: number, span: number, amountPt: number): void {
  if (amountPt <= EPSILON_PT || span <= 0) return;
  const currentPt = spanSum(widths, start, span);
  for (let offset = 0; offset < span; offset += 1) {
    const column = start + offset;
    // §17.18.87 requires all tracks in a multi-column span to be enlarged but
    // does not prescribe the distribution. Retain existing grid proportions;
    // an all-zero span has no authored ratio, so share the delta equally.
    const ratio = currentPt > EPSILON_PT
      ? (widths[column] ?? 0) / currentPt
      : 1 / span;
    widths[column] += amountPt * ratio;
  }
}

export interface TableIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
}

/**
 * Measure only the content interval contributed by a table when it is nested
 * inside another cell. This deterministically applies the ECMA-376 §17.18.87
 * content min/max algorithm to nested content. tblW/tcW remain preferred-width
 * constraints; using a nested occurrence's final fitted width as its parent's
 * content minimum would create a circular width constraint, so recursion uses
 * the intrinsic interval instead.
 *
 * Fixed-layout tables do not acquire cell-content constraints (§17.18.87).
 * They remain one fixed-width atom using their already-projected preferred
 * width; only AutoFit tables expose recursively measured content min/max.
 */
export function measureTableIntrinsicWidths(
  input: TableColumnLayoutInput,
): TableIntrinsicWidths {
  const columnCount = requiredColumnCount(input);
  if (input.layout === 'fixed') {
    const widthsPt = fixedWidths(input, columnCount);
    const widthPt = widthsPt.reduce((sum, width) => sum + width, 0);
    return Object.freeze({ minWidthPt: widthPt, maxWidthPt: widthPt });
  }
  const minimums = new Array<number>(columnCount).fill(0);
  const maximums = new Array<number>(columnCount).fill(0);
  const cells = input.rows.flatMap((row) => row.cells)
    .sort((left, right) => left.columnSpan - right.columnSpan);
  const enforce = (widths: number[], cell: TableColumnCellConstraint, requiredPt: number): void => {
    const start = Math.max(0, cell.columnStart);
    const span = Math.max(1, Math.min(cell.columnSpan, widths.length - start));
    if (span <= 0) return;
    const deficitPt = finiteNonNegative(requiredPt) - spanSum(widths, start, span);
    if (deficitPt > EPSILON_PT) growSpan(widths, start, span, deficitPt);
  };
  for (const cell of cells) {
    enforce(minimums, cell, cell.minContentWidthPt);
    enforce(maximums, cell, Math.max(cell.minContentWidthPt, cell.maxContentWidthPt));
  }
  const minWidthPt = minimums.reduce((sum, width) => sum + width, 0);
  const maxWidthPt = Math.max(
    minWidthPt,
    maximums.reduce((sum, width) => sum + width, 0),
  );
  return Object.freeze({ minWidthPt, maxWidthPt });
}

function enforceContentConstraint(
  widths: number[],
  minimums: readonly number[],
  transferFloors: readonly number[],
  maximums: readonly number[],
  cell: TableColumnCellConstraint,
): void {
  const start = Math.max(0, cell.columnStart);
  const span = Math.max(1, Math.min(cell.columnSpan, widths.length - start));
  if (span <= 0) return;
  const requiredPt = finiteNonNegative(cell.minContentWidthPt);
  const currentPt = spanSum(widths, start, span);
  if (requiredPt <= currentPt + EPSILON_PT) return;
  const maximumPt = span === 1
    ? (maximums[start] ?? requiredPt)
    : Math.max(requiredPt, finiteNonNegative(cell.maxContentWidthPt));

  // The autofit guidance permits a deficient cell to grow anywhere from its
  // minimum through maximum while first reclaiming slack from other tracks.
  // Reclaim toward maximum, then expand the table only as far as its minimum.
  const transferredPt = shrinkOutsideSpan(
    widths,
    transferFloors,
    start,
    span,
    Math.max(0, maximumPt - currentPt),
  );
  growSpan(widths, start, span, transferredPt);

  const afterPreferredTransferPt = spanSum(widths, start, span);
  if (afterPreferredTransferPt < requiredPt - EPSILON_PT) {
    // Once other tracks have reached their absolute minima, §17.4.29 allows
    // reclaiming protected dxa width to satisfy this cell's own minimum, but
    // not to enlarge it toward its optional maximum.
    const minimumTransferPt = shrinkOutsideSpan(
      widths, minimums, start, span, requiredPt - afterPreferredTransferPt,
    );
    growSpan(widths, start, span, minimumTransferPt);
    const afterMinimumTransferPt = spanSum(widths, start, span);
    if (afterMinimumTransferPt < requiredPt - EPSILON_PT) {
      growSpan(widths, start, span, requiredPt - afterMinimumTransferPt);
    }
  }
}

function noWrapPreferredFloors(
  widths: readonly number[],
  minimums: readonly number[],
  cells: readonly TableColumnCellConstraint[],
): number[] {
  const floors = [...minimums];
  for (const cell of cells) {
    if (cell.noWrap !== true || cell.preferredWidth?.kind !== 'dxa') continue;
    const start = Math.max(0, cell.columnStart);
    const span = Math.max(1, Math.min(cell.columnSpan, widths.length - start));
    const currentPt = spanSum(widths, start, span);
    if (currentPt <= EPSILON_PT) continue;
    // The authored preference survives earlier fitting; a shrunken track
    // cannot redefine the protected width (§17.4.29).
    const protectedPt = finiteNonNegative(cell.preferredWidth.value);
    for (let column = start; column < start + span; column += 1) {
      // §17.4.29 protects the span's aggregate preferred width. Retain its
      // existing grid proportions because the rule does not assign that
      // preference to individual tracks within a merged cell.
      floors[column] = Math.max(
        floors[column] ?? 0,
        widths[column]! * protectedPt / currentPt,
      );
    }
  }
  return floors;
}

/** WORD_AUTOFIT_CONTENT_COLUMN_GROWTH (table-compatibility.ts): tblGrid is
 * initial geometry (§17.4.48), not a content maximum. The measured two-track
 * controls release an unpreferred short track's excess saved width. When both
 * tracks grow, Word interpolates their content minima toward their maxima,
 * rather than sharing deficits over the saved grid. Simultaneous growth with
 * more tracks or spans is unmeasured and keeps the existing solver result. */
function growUnpreferredColumns(
  widths: number[],
  minimums: readonly number[],
  maximums: readonly number[],
  cells: readonly TableColumnCellConstraint[],
  rows: readonly TableColumnRowConstraint[],
  availableWidthPt: number,
): void {
  // A saved track without a cell has no content maximum. In particular,
  // §17.4.85/.86 wAfter/wBefore constrain omitted row intervals; treating
  // their absent content as a zero maximum would discard the specified width.
  // Only actual unpreferred cells establish eligibility. An omitted interval
  // stays excluded even when another row has an auto cell in that track.
  const eligible = widths.map(() => false);
  for (const cell of cells) {
    if (cell.preferredWidth === null && cell.columnSpan === 1
      && cell.columnStart >= 0 && cell.columnStart < widths.length) {
      eligible[cell.columnStart] = true;
    }
  }
  for (const cell of cells) {
    if (cell.preferredWidth === null && cell.columnSpan === 1) continue;
    const start = Math.max(0, cell.columnStart);
    const end = Math.min(widths.length, start + Math.max(1, cell.columnSpan));
    for (let column = start; column < end; column += 1) eligible[column] = false;
  }
  // All omissions are prefixes/suffixes. Union them before walking tracks,
  // avoiding a repeated full-grid scan for each row.
  let beforeSpan = 0;
  let afterSpan = 0;
  for (const row of rows) {
    beforeSpan = Math.max(beforeSpan, row.before?.columnSpan ?? 0);
    afterSpan = Math.max(afterSpan, row.after?.columnSpan ?? 0);
  }
  for (let column = 0; column < Math.min(widths.length, beforeSpan); column += 1) eligible[column] = false;
  for (let column = Math.max(0, widths.length - afterSpan); column < widths.length; column += 1) eligible[column] = false;
  const twoUnpreferredTracks = widths.length === 2 && eligible.every(Boolean);
  if (twoUnpreferredTracks) {
    widths.forEach((width, column) => {
      widths[column] = Math.min(width, maximums[column] ?? width);
    });
  }
  const deficits = widths.map((width, column) => (eligible[column]
    ? Math.max(0, (maximums[column] ?? 0) - width)
    : 0));
  const totalDeficitPt = deficits.reduce((sum, deficit) => sum + deficit, 0);
  if (totalDeficitPt <= EPSILON_PT) return;
  if (deficits.filter((deficit) => deficit > EPSILON_PT).length > 1) {
    if (!twoUnpreferredTracks) return;
    const minimumPt = minimums.reduce((sum, width) => sum + width, 0);
    if (minimumPt > availableWidthPt) return;
    const intervals = maximums.map((maximum, column) => Math.max(0, maximum - (minimums[column] ?? 0)));
    const totalIntervalPt = intervals.reduce((sum, width) => sum + width, 0);
    if (totalIntervalPt <= EPSILON_PT) return;
    const fraction = Math.min(1, (availableWidthPt - minimumPt) / totalIntervalPt);
    widths.forEach((_width, column) => {
      widths[column] = (minimums[column] ?? 0) + fraction * intervals[column]!;
    });
    return;
  }
  const roomPt = availableWidthPt - widths.reduce((sum, width) => sum + width, 0);
  if (roomPt <= EPSILON_PT) return;
  const growPt = Math.min(roomPt, totalDeficitPt);
  deficits.forEach((deficit, column) => {
    widths[column] = (widths[column] ?? 0) + growPt * deficit / totalDeficitPt;
  });
}

function fitToAvailableWidth(
  widths: number[],
  minimums: readonly number[],
  cells: readonly TableColumnCellConstraint[],
  availableWidthPt: number,
  outerMarginAllowancePt: number,
): number[] {
  const totalPt = widths.reduce((sum, width) => sum + width, 0);
  if (totalPt <= availableWidthPt + EPSILON_PT || totalPt <= EPSILON_PT) return widths;

  const result = [...widths];
  // ECMA-376 §17.4.29: a noWrap dxa cell retains its preferred width while
  // other cells can still shrink to their absolute content minima. A later
  // pass may shrink that preference if those other minima exhaust the band.
  const protectedFloors = noWrapPreferredFloors(result, minimums, cells);
  const shrinkToFloors = (floors: readonly number[], requestedPt: number): number => {
    const slack = result.map((width, column) => Math.max(0, width - (floors[column] ?? 0)));
    const totalSlackPt = slack.reduce((sum, value) => sum + value, 0);
    const shrinkPt = Math.min(requestedPt, totalSlackPt);
    if (shrinkPt <= EPSILON_PT || totalSlackPt <= EPSILON_PT) return 0;
    result.forEach((_width, column) => {
      result[column] -= shrinkPt * ((slack[column] ?? 0) / totalSlackPt);
    });
    return shrinkPt;
  };
  const requestedShrinkPt = totalPt - availableWidthPt;
  const protectedShrinkPt = shrinkToFloors(protectedFloors, requestedShrinkPt);
  shrinkToFloors(minimums, requestedShrinkPt - protectedShrinkPt);

  // A proportional column shrink can violate an otherwise satisfiable minimum
  // on a spanning cell even though another track still has slack. Restore each
  // spanning minimum by transferring that outside slack before declaring that
  // forced line breaks are necessary.
  for (const cell of cells) {
    if (cell.columnSpan <= 1) continue;
    const start = Math.max(0, cell.columnStart);
    const span = Math.max(1, Math.min(cell.columnSpan, result.length - start));
    const deficitPt = finiteNonNegative(cell.minContentWidthPt) - spanSum(result, start, span);
    if (deficitPt <= EPSILON_PT) continue;
    const preferredTransferPt = shrinkOutsideSpan(result, protectedFloors, start, span, deficitPt);
    const transferredPt = preferredTransferPt + shrinkOutsideSpan(
      result, minimums, start, span, deficitPt - preferredTransferPt,
    );
    growSpan(result, start, span, transferredPt);
    if (transferredPt < deficitPt - EPSILON_PT) {
      growSpan(result, start, span, deficitPt - transferredPt);
    }
  }

  const afterSlackPt = result.reduce((sum, width) => sum + width, 0);
  if (afterSlackPt <= availableWidthPt + EPSILON_PT || afterSlackPt <= EPSILON_PT) {
    return cleanWidths(result);
  }
  const autoNoWrapCell = cells.find((cell) => cell.noWrap === true
    && cell.preferredWidth === null && cell.columnSpan === 1);
  if (
    outerMarginAllowancePt > EPSILON_PT
    && (autoNoWrapCell?.minContentWidthPt ?? 0)
      > availableWidthPt - outerMarginAllowancePt + EPSILON_PT
    && result.length === 2
    && cells.length === 2
    && cells.every((cell) => cell.columnSpan === 1)
    && cells.some((cell) => cell.columnStart === 0)
    && cells.some((cell) => cell.columnStart === 1)
    && autoNoWrapCell
    && cells.filter((cell) => cell.noWrap === true).length === 1
  ) {
    // WORD_AUTOFIT_NOWRAP_AUTO_FORCED_FIT (table-compatibility.ts): scale
    // content-only widths against the text band, then distribute the outer
    // margin allowance according to the OTHER cell's content width. §17.18.87
    // does not define the forced-break distribution; the recorded evidence in
    // that rule is limited to this two-cell case.
    const margins = [0, 0];
    for (const cell of cells) margins[cell.columnStart] = finiteNonNegative(cell.horizontalMarginsPt);
    const contentWidths = result.map((width, column) => Math.max(0, width - (margins[column] ?? 0)));
    const contentTotalPt = contentWidths.reduce((sum, width) => sum + width, 0);
    if (contentTotalPt > EPSILON_PT) {
      const contentBandPt = Math.max(0, availableWidthPt - outerMarginAllowancePt);
      const scaled = contentWidths.map((width) => width * contentBandPt / contentTotalPt);
      const noWrapColumn = autoNoWrapCell.columnStart;
      const ordinaryColumn = 1 - noWrapColumn;
      scaled[noWrapColumn] += outerMarginAllowancePt
        * (contentWidths[ordinaryColumn] ?? 0) / contentTotalPt;
      scaled[ordinaryColumn] += outerMarginAllowancePt
        * (contentWidths[noWrapColumn] ?? 0) / contentTotalPt;
      return cleanWidths(scaled);
    }
  }
  // The jointly retained minima do not fit, so §17.18.87 now permits forced
  // line breaks. It does not prescribe the final track distribution; reuse the
  // proportional reduction defined by its fixed-layout guidance as an explicit
  // deterministic solver policy, not as a compatibility constant.
  const scale = Math.max(0, availableWidthPt) / afterSlackPt;
  return cleanWidths(result.map((width) => width * scale));
}

/**
 * ECMA-376 §17.18.87 fixed/autofit guidance expressed as a deterministic,
 * parser-independent constraint solver. `tblGrid` seeds widths; it never
 * suppresses authored row/cell preferences.
 */
function solveTableColumnWidths(input: TableColumnLayoutInput): readonly number[] {
  const columnCount = requiredColumnCount(input);
  if (columnCount === 0) return Object.freeze([]);
  const widths = fixedWidths(input, columnCount);
  if (input.layout === 'fixed') {
    if (input.availableWidthPt === null) {
      return Object.freeze(widths);
    }
    const totalPt = widths.reduce((sum, width) => sum + width, 0);
    const maximumPt = finiteNonNegative(input.availableWidthPt);
    if (totalPt <= maximumPt + EPSILON_PT || totalPt <= EPSILON_PT) {
      return Object.freeze(widths);
    }
    // The caller supplies the physical width ceiling for this table occurrence.
    // §17.18.87 itself reduces fixed tracks to the preferred table width; the
    // containing page/story projection owns any additional physical ceiling.
    const scale = maximumPt / totalPt;
    return Object.freeze(cleanWidths(widths.map((width) => width * scale)));
  }

  const percentageBasePt = widths.reduce((sum, width) => sum + width, 0);
  const { minimums, maximums } = singleColumnBounds(
    input.rows,
    columnCount,
    percentageBasePt,
  );
  const cells = input.rows.flatMap((row) => row.cells);
  // Single-column min/max bounds are established before spanning constraints.
  cells.sort((left, right) => left.columnSpan - right.columnSpan);
  const transferFloors = noWrapPreferredFloors(widths, minimums, cells);
  for (const cell of cells) {
    enforceContentConstraint(widths, minimums, transferFloors, maximums, cell);
  }
  const totalMaximumPt = maximums.reduce((sum, width) => sum + width, 0);
  const nonbinding = totalMaximumPt <= finiteNonNegative(input.availableWidthPt) + EPSILON_PT
    && input.contentBandPt !== undefined && totalMaximumPt < finiteNonNegative(input.contentBandPt) - EPSILON_PT;
  if ((input.growUnpreferredColumns === true || (input.growUnpreferredColumns === 'nonbinding' && nonbinding))
    && input.tablePreferredWidthPt === null) {
    growUnpreferredColumns(widths, minimums, maximums, cells, input.rows, finiteNonNegative(input.availableWidthPt));
  }
  return Object.freeze(fitToAvailableWidth(
    widths,
    minimums,
    cells,
    finiteNonNegative(input.availableWidthPt),
    finiteNonNegative(input.outerMarginAllowancePt),
  ));
}

export interface ResolvedTableColumnLayout {
  readonly widthsPt: readonly number[];
  readonly widthKeys: readonly (ExactLengthKey | null)[];
}

/** Preserve an unchanged authored track's identity state exactly. A hand-built
 * numeric track has no authored key array and therefore receives its numeric
 * definition; a solver-changed track likewise receives its final numeric
 * definition. Authored `null` is different: it means the exact identity is
 * unavailable and must remain unknown while its geometry is unchanged. */
export function resolveTableColumnLayout(input: TableColumnLayoutInput): ResolvedTableColumnLayout {
  const widthsPt = solveTableColumnWidths(input);
  const widthKeys = widthsPt.map((width, column) => {
    const changed = width !== input.gridWidthsPt[column];
    if (!changed && input.gridWidthKeys?.[column] === null) return null;
    if (!changed && input.gridWidthKeys?.[column] !== undefined) {
      return input.gridWidthKeys[column]!;
    }
    return exactLengthKeyFromNumber(width) ?? '0/1';
  });
  return Object.freeze({
    widthsPt: Object.freeze([...widthsPt]),
    widthKeys: Object.freeze(widthKeys),
  });
}

export function resolveTableColumnWidths(input: TableColumnLayoutInput): readonly number[] {
  return resolveTableColumnLayout(input).widthsPt;
}
