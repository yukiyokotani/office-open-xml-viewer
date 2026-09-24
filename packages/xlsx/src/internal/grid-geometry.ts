import type { ViewportRange, Worksheet } from '../types.js';
import { baseColWidthToPx, colWidthToPx, rowHeightToPx } from './grid-metrics.js';
import { isMacDesktop } from './platform.js';
import { GridAxisGeometry } from './grid-axis-geometry.js';

export { GridAxisGeometry } from './grid-axis-geometry.js';

export const MAX_WORKSHEET_ROW = 1_048_576;
export const MAX_WORKSHEET_COL = 16_384;

export interface GridCellRectOptions {
  readonly scale: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly headerWidth: number;
  readonly headerHeight: number;
}

export interface GridVisibleRangeOptions extends GridCellRectOptions {
  readonly width: number;
  readonly height: number;
  readonly buffer?: number;
}

export interface GridVisibleGeometry {
  readonly range: ViewportRange;
  readonly offsetX: number;
  readonly offsetY: number;
  readonly frozenWidth: number;
  readonly frozenHeight: number;
}

export interface GridScrollToCellOptions {
  readonly scale: number;
  readonly viewportWidth: number;
  readonly viewportHeight: number;
  readonly currentX: number;
  readonly currentY: number;
  readonly headerWidth: number;
  readonly headerHeight: number;
  readonly align: 'nearest' | 'start' | 'center' | 'end';
}

/** Pure worksheet geometry shared by both XLSX viewer facades. */
export class GridGeometry {
  private static readonly cache = new WeakMap<Worksheet, {
    readonly mdw: number;
    readonly geometry: GridGeometry;
  }>();

  static forWorksheet(worksheet: Worksheet, mdw: number): GridGeometry {
    const cached = this.cache.get(worksheet);
    if (cached && Object.is(cached.mdw, mdw)) return cached.geometry;
    const geometry = new GridGeometry(worksheet, mdw);
    this.cache.set(worksheet, { mdw, geometry });
    return geometry;
  }

  /** Resolve MDW only when this worksheet has no live geometry snapshot. */
  static forWorksheetMeasured(
    worksheet: Worksheet,
    measureMdw: () => number,
  ): GridGeometry {
    const cached = this.cache.get(worksheet);
    if (cached) return cached.geometry;
    return this.forWorksheet(worksheet, measureMdw());
  }

  static invalidate(worksheet: Worksheet): void {
    this.cache.delete(worksheet);
  }

  readonly col: GridAxisGeometry;
  readonly row: GridAxisGeometry;
  readonly maximumDigitWidth: number;
  private readonly freezeRows: number;
  private readonly freezeCols: number;
  private scaledCache: Readonly<{
    scale: number;
    row: GridAxisGeometry;
    col: GridAxisGeometry;
  }> | null = null;

  private constructor(worksheet: Worksheet, mdw: number) {
    this.maximumDigitWidth = mdw;
    this.freezeRows = Math.min(MAX_WORKSHEET_ROW, Math.max(0, worksheet.freezeRows ?? 0));
    this.freezeCols = Math.min(MAX_WORKSHEET_COL, Math.max(0, worksheet.freezeCols ?? 0));
    const defaultColPx = worksheet.baseColWidth === undefined
      ? colWidthToPx(worksheet.defaultColWidth, mdw)
      : baseColWidthToPx(worksheet.baseColWidth, mdw, isMacDesktop());
    const resolvedColWidths = new Float64Array(MAX_WORKSHEET_COL + 1);
    resolvedColWidths.fill(Number.NaN);
    // Resolve declarations from last to first. The disjoint-set successor
    // skips columns already assigned by a later declaration, so every sheet
    // column is written at most once even when an adversarial file contains
    // many overlapping full-width ranges.
    const successor = new Int32Array(MAX_WORKSHEET_COL + 2);
    for (let index = 1; index < successor.length; index++) successor[index] = index;
    const findUnassigned = (start: number): number => {
      let root = start;
      while (successor[root] !== root) root = successor[root];
      let index = start;
      while (successor[index] !== index) {
        const next = successor[index];
        successor[index] = root;
        index = next;
      }
      return root;
    };
    const authoredRanges = worksheet.colWidthRanges ?? [];
    for (let rangeIndex = authoredRanges.length - 1; rangeIndex >= 0; rangeIndex--) {
      const range = authoredRanges[rangeIndex];
      const min = Math.max(1, Math.min(MAX_WORKSHEET_COL, Math.trunc(range.min)));
      const max = Math.max(1, Math.min(MAX_WORKSHEET_COL, Math.trunc(range.max)));
      if (max < min || !Number.isFinite(range.width) || range.width < 0) continue;
      let index = findUnassigned(min);
      while (index <= max) {
        resolvedColWidths[index] = range.width;
        successor[index] = findUnassigned(index + 1);
        index = successor[index];
      }
    }
    // Explicit point widths are also used by live resize overrides, so they
    // intentionally win over the parsed compact ranges.
    for (const [rawIndex, width] of Object.entries(worksheet.colWidths)) {
      const index = Number(rawIndex);
      if (Number.isInteger(index) && index >= 1 && index <= MAX_WORKSHEET_COL) {
        resolvedColWidths[index] = width;
      }
    }
    const columnPixels: Array<{ index: number; px: number }> = [];
    for (let index = 1; index <= MAX_WORKSHEET_COL; index++) {
      const width = resolvedColWidths[index];
      if (!Number.isFinite(width) || width < 0) continue;
      const px = colWidthToPx(width, mdw);
      if (px !== defaultColPx) columnPixels.push({ index, px });
    }
    this.col = new GridAxisGeometry(
      {},
      defaultColPx,
      (raw) => colWidthToPx(raw, mdw),
      MAX_WORKSHEET_COL,
      columnPixels,
    );
    this.row = new GridAxisGeometry(
      worksheet.rowHeights,
      rowHeightToPx(worksheet.defaultRowHeight),
      rowHeightToPx,
      MAX_WORKSHEET_ROW,
    );
  }

  logicalFrozenExtent(): { width: number; height: number } {
    return {
      width: this.col.offsetOf(this.freezeCols + 1),
      height: this.row.offsetOf(this.freezeRows + 1),
    };
  }

  roundedFrozenExtent(scale: number): { width: number; height: number } {
    const axes = this.axesAtScale(scale);
    return {
      width: axes.col.offsetOf(this.freezeCols + 1),
      height: axes.row.offsetOf(this.freezeRows + 1),
    };
  }

  /** Frozen bands that can physically reach a viewport at this scale. */
  effectiveFrozenBands(options: {
    readonly scale: number;
    readonly width: number;
    readonly height: number;
    readonly headerWidth: number;
    readonly headerHeight: number;
    readonly rows: number;
    readonly cols: number;
  }): { rows: number; cols: number } {
    const axes = this.axesAtScale(options.scale);
    const rowBands = axes.row.bandsToCover(
      1,
      Math.max(0, options.rows),
      Math.max(0, options.height - Math.round(options.headerHeight * options.scale)),
    );
    const colBands = axes.col.bandsToCover(
      1,
      Math.max(0, options.cols),
      Math.max(0, options.width - Math.round(options.headerWidth * options.scale)),
    );
    return {
      rows: rowBands.at(-1)?.index ?? 0,
      cols: colBands.at(-1)?.index ?? 0,
    };
  }

  logicalContentExtent(
    maxRow: number,
    maxCol: number,
    headerWidth: number,
    headerHeight: number,
  ): { width: number; height: number } {
    return {
      width: headerWidth + this.col.offsetOf(Math.min(MAX_WORKSHEET_COL, maxCol) + 1),
      height: headerHeight + this.row.offsetOf(Math.min(MAX_WORKSHEET_ROW, maxRow) + 1),
    };
  }

  roundedContentExtent(
    maxRow: number,
    maxCol: number,
    scale: number,
    headerWidth: number,
    headerHeight: number,
  ): { width: number; height: number } {
    const axes = this.axesAtScale(scale);
    return {
      width: Math.round(headerWidth * scale)
        + axes.col.offsetOf(Math.min(MAX_WORKSHEET_COL, maxCol) + 1),
      height: Math.round(headerHeight * scale)
        + axes.row.offsetOf(Math.min(MAX_WORKSHEET_ROW, maxRow) + 1),
    };
  }

  cellAt(
    innerX: number,
    innerY: number,
    viewport: { readonly scrollX: number; readonly scrollY: number; readonly scale?: number },
  ): { row: number; col: number } | null {
    if (innerX < 0 || innerY < 0) return null;
    const row = this.rowAt(innerY, viewport.scrollY, viewport.scale);
    if (row === null) return null;
    const col = this.colAt(innerX, viewport.scrollX, viewport.scale);
    return col === null ? null : { row, col };
  }

  rowAt(innerY: number, scrollY: number, scale = 1): number | null {
    if (innerY < 0) return null;
    const row = this.axesAtScale(scale).row;
    const frozenHeight = row.offsetOf(this.freezeRows + 1);
    return innerY < frozenHeight
      ? this.indexWithinFrozen(row, innerY, this.freezeRows)
      : row.scrollableIndexAt(innerY - frozenHeight + scrollY, this.freezeRows + 1);
  }

  colAt(innerX: number, scrollX: number, scale = 1): number | null {
    if (innerX < 0) return null;
    const col = this.axesAtScale(scale).col;
    const frozenWidth = col.offsetOf(this.freezeCols + 1);
    return innerX < frozenWidth
      ? this.indexWithinFrozen(col, innerX, this.freezeCols)
      : col.scrollableIndexAt(innerX - frozenWidth + scrollX, this.freezeCols + 1);
  }

  cellRect(
    row: number,
    col: number,
    options: GridCellRectOptions,
  ): { x: number; y: number; w: number; h: number } | null {
    if (row < 1 || row > MAX_WORKSHEET_ROW || col < 1 || col > MAX_WORKSHEET_COL) return null;
    const axes = this.axesAtScale(options.scale);
    const headerX = Math.round(options.headerWidth * options.scale);
    const headerY = Math.round(options.headerHeight * options.scale);
    const frozen = this.roundedFrozenExtent(options.scale);

    const x = col <= this.freezeCols
      ? headerX + axes.col.offsetOf(col)
      : this.scrollableCellPosition(
          axes.col,
          col,
          this.freezeCols,
          options.scrollX,
          headerX + frozen.width,
        );
    const y = row <= this.freezeRows
      ? headerY + axes.row.offsetOf(row)
      : this.scrollableCellPosition(
          axes.row,
          row,
          this.freezeRows,
          options.scrollY,
          headerY + frozen.height,
        );
    return { x, y, w: axes.col.sizeOf(col), h: axes.row.sizeOf(row) };
  }

  visibleRange(options: GridVisibleRangeOptions): GridVisibleGeometry {
    const axes = this.axesAtScale(options.scale);
    const frozen = this.roundedFrozenExtent(options.scale);
    const colStart = axes.col.indexAt(options.scrollX + axes.col.offsetOf(this.freezeCols + 1));
    const rowStart = axes.row.indexAt(options.scrollY + axes.row.offsetOf(this.freezeRows + 1));
    const cellWidth = options.width - Math.round(options.headerWidth * options.scale) - frozen.width;
    const cellHeight = options.height - Math.round(options.headerHeight * options.scale) - frozen.height;
    const buffer = options.buffer ?? 0;
    return {
      range: {
        row: rowStart.index,
        col: colStart.index,
        rows: axes.row.countToCover(rowStart.index, cellHeight + rowStart.partial * 2) + buffer,
        cols: axes.col.countToCover(colStart.index, cellWidth + colStart.partial * 2) + buffer,
      },
      offsetX: colStart.partial / options.scale,
      offsetY: rowStart.partial / options.scale,
      frozenWidth: frozen.width / options.scale,
      frozenHeight: frozen.height / options.scale,
    };
  }

  scrollOffsetForCell(
    row: number,
    col: number,
    options: GridScrollToCellOptions,
  ): { x: number; y: number } {
    const scaledFrozen = this.roundedFrozenExtent(options.scale);
    const viewTop = Math.round(options.headerHeight * options.scale) + scaledFrozen.height;
    const viewLeft = Math.round(options.headerWidth * options.scale) + scaledFrozen.width;
    let y = options.currentY;
    const axes = this.axesAtScale(options.scale);
    if (row > this.freezeRows && row <= MAX_WORKSHEET_ROW) {
      const cellStart = axes.row.offsetOf(row) - axes.row.offsetOf(this.freezeRows + 1);
      const cellSize = axes.row.sizeOf(row);
      y = this.alignedOffset(
        cellStart,
        cellSize,
        options.currentY,
        viewTop,
        options.viewportHeight,
        options.align,
      );
    }
    let x = options.currentX;
    if (col > this.freezeCols && col <= MAX_WORKSHEET_COL) {
      const cellStart = axes.col.offsetOf(col) - axes.col.offsetOf(this.freezeCols + 1);
      const cellSize = axes.col.sizeOf(col);
      x = this.alignedOffset(
        cellStart,
        cellSize,
        options.currentX,
        viewLeft,
        options.viewportWidth,
        options.align,
      );
    }
    return { x: Math.max(0, x), y: Math.max(0, y) };
  }

  /** Rounded CSS-pixel axes used by paint, hit testing, and anchored objects. */
  axesAtScale(scale: number): Readonly<{ row: GridAxisGeometry; col: GridAxisGeometry }> {
    if (this.scaledCache?.scale === scale) return this.scaledCache;
    const scaled = { scale, row: this.row.scaled(scale), col: this.col.scaled(scale) };
    this.scaledCache = scaled;
    return scaled;
  }

  private indexWithinFrozen(
    axis: GridAxisGeometry,
    offset: number,
    frozenCount: number,
  ): number | null {
    if (frozenCount === 0) return null;
    const index = axis.indexAt(offset).index;
    return index <= frozenCount ? index : null;
  }

  private scrollableCellPosition(
    scaledAxis: GridAxisGeometry,
    index: number,
    frozenCount: number,
    scroll: number,
    scrollAreaStart: number,
  ): number {
    const start = scaledAxis.indexAt(scroll + scaledAxis.offsetOf(frozenCount + 1));
    return scrollAreaStart - start.partial
      + scaledAxis.offsetOf(index) - scaledAxis.offsetOf(start.index);
  }

  private alignedOffset(
    cellStart: number,
    cellSize: number,
    current: number,
    viewportStart: number,
    viewportEnd: number,
    align: GridScrollToCellOptions['align'],
  ): number {
    const cellOnScreen = viewportStart + cellStart - current;
    if (align === 'start') return cellStart;
    if (align === 'center') return cellStart - (viewportEnd - viewportStart - cellSize) / 2;
    if (align === 'end') return cellStart - (viewportEnd - viewportStart - cellSize);
    if (cellOnScreen < viewportStart) return cellStart;
    if (cellOnScreen + cellSize > viewportEnd) {
      return cellStart - (viewportEnd - viewportStart - cellSize);
    }
    return current;
  }
}
