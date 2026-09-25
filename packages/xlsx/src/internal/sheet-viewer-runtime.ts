import type { Worksheet } from '../types.js';
import type { XlsxWorkbook } from '../workbook.js';
import type { CellAddress, XlsxSelectionArea, XlsxSelectionState } from '../selection.js';
import { areaContainsCell, MAX_SELECTION_AREAS, normalizeSelectionState } from '../selection.js';
import {
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';

/** Generation-safe workbook ownership for one viewer instance. */
export class SheetAcquisition {
  private readonly owner = new TerminalResourceOwner<XlsxWorkbook>('SheetAcquisition');

  get current(): XlsxWorkbook | null {
    return this.owner.current;
  }

  async replace(
    load: () => Promise<XlsxWorkbook>,
    beforeCommit?: (previous: XlsxWorkbook | null) => void,
  ): Promise<XlsxWorkbook | null> {
    return await this.owner.replace(load, beforeCommit);
  }

  /** Commit an already acquired workbook, closing the previously owned one.
   *  Pass `owned: false` when the caller retains workbook lifecycle ownership. */
  install(candidate: XlsxWorkbook, owned = true): void {
    this.owner.install(candidate, owned);
  }

  destroy(): void {
    this.owner.close();
  }
}

/** Create the viewer-owned mutable projection of one cached worksheet.
 * Cell/content graphs remain shared and read-only; only the maps and row flags
 * changed by view-only resize/outline interactions are copied. */
export function createSheetViewModel(source: Worksheet): Worksheet {
  return {
    ...source,
    rows: source.rows.map((row) => ({ ...row })),
    rowHeights: { ...source.rowHeights },
    colWidths: { ...source.colWidths },
    colCollapsed: source.colCollapsed ? { ...source.colCollapsed } : undefined,
  };
}

/** Logical viewport independent of native browser scrolling. */
export class ViewportState {
  private offsetX = 0;
  private offsetY = 0;
  private contentWidth = 0;
  private contentHeight = 0;
  private viewportWidth = 0;
  private viewportHeight = 0;

  constructor(private scaleValue: number) {}

  get x(): number { return this.offsetX; }
  get y(): number { return this.offsetY; }
  get scale(): number { return this.scaleValue; }
  get maxX(): number { return Math.max(0, this.contentWidth - this.viewportWidth); }
  get maxY(): number { return Math.max(0, this.contentHeight - this.viewportHeight); }

  setScale(scale: number): void {
    this.scaleValue = scale;
  }

  setExtent(contentWidth: number, contentHeight: number): void {
    this.contentWidth = Math.max(0, contentWidth);
    this.contentHeight = Math.max(0, contentHeight);
    this.setOffset(this.offsetX, this.offsetY);
  }

  /** Expand to include an extent reported by a native scroll host without
   * shrinking the format-computed worksheet extent. Browser layout and test
   * doubles may publish scrollWidth/scrollHeight after the worksheet geometry
   * has already been installed. */
  ensureExtent(contentWidth: number, contentHeight: number): void {
    this.contentWidth = Math.max(this.contentWidth, Math.max(0, contentWidth));
    this.contentHeight = Math.max(this.contentHeight, Math.max(0, contentHeight));
  }

  setViewportSize(width: number, height: number): void {
    this.viewportWidth = Math.max(0, width);
    this.viewportHeight = Math.max(0, height);
    this.setOffset(this.offsetX, this.offsetY);
  }

  setOffset(x: number, y: number): void {
    this.offsetX = Math.min(this.maxX, Math.max(0, x));
    this.offsetY = Math.min(this.maxY, Math.max(0, y));
  }

  /** Mirror offsets already clamped by a native browser scroll container. This
   * intentionally does not re-clamp against logical extents: the DOM is the
   * authority for the composite viewer and may publish its geometry later. */
  adoptNativeOffset(x: number, y: number): void {
    this.offsetX = Math.max(0, x);
    this.offsetY = Math.max(0, y);
  }

  reset(): void {
    this.offsetX = 0;
    this.offsetY = 0;
  }
}

/** XLSX render scheduling around core-owned static bitmap lifecycle mechanics. */
export class SheetRenderDispatcher {
  private animationFrame: number | null = null;
  private activeRender = false;
  private pendingRender: (() => void | Promise<void>) | null = null;
  private readonly staticDispatcher: StaticCanvasRenderDispatcher | null;
  private readonly frameScheduler: Pick<Window, 'requestAnimationFrame' | 'cancelAnimationFrame'> | null;
  private generation = 0;
  private destroyed = false;

  constructor(
    canvas?: HTMLCanvasElement,
    workerBitmapMode = false,
    frameScheduler?: Partial<Pick<Window, 'requestAnimationFrame' | 'cancelAnimationFrame'>> | null,
  ) {
    const scheduler = frameScheduler ?? globalThis;
    this.frameScheduler =
      typeof scheduler.requestAnimationFrame === 'function' &&
      typeof scheduler.cancelAnimationFrame === 'function'
        ? {
            requestAnimationFrame: (callback) => scheduler.requestAnimationFrame!(callback),
            cancelAnimationFrame: (handle) => scheduler.cancelAnimationFrame!(handle),
          }
        : null;
    this.staticDispatcher = canvas
      ? new StaticCanvasRenderDispatcher(canvas, workerBitmapMode)
      : null;
  }

  begin(): number {
    if (this.staticDispatcher) return this.staticDispatcher.begin();
    return ++this.generation;
  }

  isCurrent(generation: number): boolean {
    if (this.staticDispatcher) return this.staticDispatcher.isCurrent(generation);
    return !this.destroyed && generation === this.generation;
  }

  /** Delegate stale disposal and atomic bitmap replacement to the core owner. */
  commitBitmap(
    generation: number,
    bitmap: ImageBitmap,
    cssWidth: number,
    cssHeight: number,
  ): boolean {
    if (!this.isCurrent(generation)) {
      bitmap.close();
      return false;
    }
    if (!this.staticDispatcher) {
      bitmap.close();
      throw new Error('SheetRenderDispatcher is not configured for worker bitmap rendering');
    }
    return this.staticDispatcher.commitBitmap(generation, bitmap, {
      cssWidth,
      cssHeight,
    });
  }

  schedule(render: () => void | Promise<void>): void {
    if (this.destroyed) return;
    this.pendingRender = render;
    if (this.activeRender) {
      // A queued viewport supersedes the frame currently awaiting a worker
      // bitmap immediately, not only when the queued callback eventually gets
      // its backpressure slot. Otherwise the old bitmap can still commit after
      // scroll/resize state has changed and briefly disagree with the live
      // gutters/overlays. The queued callback calls begin() again when it starts
      // and becomes the sole current generation.
      this.begin();
      return;
    }
    if (this.animationFrame !== null) return;
    this.queuePendingRender();
  }

  private queuePendingRender(): void {
    if (!this.frameScheduler) {
      this.startPendingRender();
      return;
    }
    this.animationFrame = this.frameScheduler.requestAnimationFrame(() => {
      this.animationFrame = null;
      this.startPendingRender();
    });
  }

  private startPendingRender(): void {
    if (this.destroyed || this.activeRender) return;
    const render = this.pendingRender;
    this.pendingRender = null;
    if (!render) return;
    this.activeRender = true;
    let completion: void | Promise<void>;
    try {
      completion = render();
    } catch {
      completion = undefined;
    }
    Promise.resolve(completion)
      // Scheduling has no returned promise. Callers that need error delivery
      // route it inside `render`; keep a thrown callback from becoming an
      // unhandled rejection while still releasing the backpressure slot.
      .catch(() => undefined)
      .finally(() => {
        this.activeRender = false;
        if (!this.destroyed && this.pendingRender) this.queuePendingRender();
      });
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.pendingRender = null;
    if (this.animationFrame !== null && this.frameScheduler) {
      this.frameScheduler.cancelAnimationFrame(this.animationFrame);
      this.animationFrame = null;
    }
    this.staticDispatcher?.destroy();
    this.generation++;
  }
}

export type SheetSelectionMode = 'cells' | 'rows' | 'cols' | 'all';

/** Selection state and immutable snapshots for one sheet viewer instance. */
export class SelectionController {
  private state: XlsxSelectionState | null = null;
  private dragPointerId: number | null = null;

  get anchor(): CellAddress | null {
    return this.state ? { ...this.state.extensionAnchor } : null;
  }

  get active(): CellAddress | null {
    return this.state ? { ...this.state.activeCell } : null;
  }

  get mode(): SheetSelectionMode {
    const area = this.activeArea;
    if (!area) return 'cells';
    return area.kind === 'columns' ? 'cols' : area.kind === 'sheet' ? 'all' : area.kind;
  }
  get dragging(): boolean { return this.dragPointerId !== null; }
  get draggingPointerId(): number | null { return this.dragPointerId; }

  get activeArea(): XlsxSelectionArea | null {
    return this.state?.areas[this.state.activeAreaIndex] ?? null;
  }

  beginDrag(pointerId: number): void {
    this.dragPointerId = pointerId;
  }

  endDrag(pointerId?: number): void {
    if (pointerId === undefined || pointerId === this.dragPointerId) this.dragPointerId = null;
  }

  reset(): void {
    this.state = null;
    this.dragPointerId = null;
  }

  setState(state: XlsxSelectionState | null): void {
    this.state = state ? normalizeSelectionState(state) : null;
  }

  select(cell: CellAddress, mode: SheetSelectionMode = 'cells'): void {
    const area: XlsxSelectionArea = mode === 'rows'
      ? { kind: 'rows', firstRow: cell.row, lastRow: cell.row }
      : mode === 'cols'
        ? { kind: 'columns', firstColumn: cell.col, lastColumn: cell.col }
        : mode === 'all'
          ? { kind: 'sheet' }
          : { kind: 'cells', top: cell.row, left: cell.col, bottom: cell.row, right: cell.col };
    this.state = normalizeSelectionState({
      areas: [area], activeAreaIndex: 0, activeCell: cell, extensionAnchor: cell,
    });
  }

  /** Add one Ctrl/Cmd-started selection area and make it active. A coordinate
   * already covered by a same-kind area activates that area without duplicating
   * it. Subsequent drag/Shift extension changes only a newly appended range.
   * Returns whether pointer drag may extend a new area. */
  add(cell: CellAddress, mode: SheetSelectionMode = 'cells'): boolean {
    if (!this.state) {
      this.select(cell, mode);
      return true;
    }
    const existingIndex = this.state.areas.findIndex((area) => {
      const sameMode = mode === 'rows'
        ? area.kind === 'rows'
        : mode === 'cols'
          ? area.kind === 'columns'
          : mode === 'all'
            ? area.kind === 'sheet'
            : area.kind === 'cells';
      return sameMode && areaContainsCell(area, cell);
    });
    if (existingIndex >= 0) {
      // A selected coordinate is already represented by this area. Keep one
      // canonical area instead of appending duplicates until the resource cap;
      // make it active so activeCell/activeAreaIndex still follow the pointer.
      this.state = normalizeSelectionState({
        ...this.state,
        activeAreaIndex: existingIndex,
        activeCell: cell,
        extensionAnchor: cell,
      });
      return false;
    }
    if (this.state.areas.length >= MAX_SELECTION_AREAS) return false;
    const area: XlsxSelectionArea = mode === 'rows'
      ? { kind: 'rows', firstRow: cell.row, lastRow: cell.row }
      : mode === 'cols'
        ? { kind: 'columns', firstColumn: cell.col, lastColumn: cell.col }
        : mode === 'all'
          ? { kind: 'sheet' }
          : { kind: 'cells', top: cell.row, left: cell.col, bottom: cell.row, right: cell.col };
    const areas = [...this.state.areas, area];
    this.state = normalizeSelectionState({
      areas,
      activeAreaIndex: areas.length - 1,
      activeCell: cell,
      extensionAnchor: cell,
    });
    return true;
  }

  extend(cell: CellAddress): void {
    if (!this.state) { this.select(cell); return; }
    const anchor = this.state.extensionAnchor;
    const area = this.activeArea;
    if (!area) return;
    const extended: XlsxSelectionArea = area.kind === 'rows'
      ? { kind: 'rows', firstRow: Math.min(anchor.row, cell.row), lastRow: Math.max(anchor.row, cell.row) }
      : area.kind === 'columns'
        ? { kind: 'columns', firstColumn: Math.min(anchor.col, cell.col), lastColumn: Math.max(anchor.col, cell.col) }
        : area.kind === 'sheet'
          ? area
          : {
              kind: 'cells',
              top: Math.min(anchor.row, cell.row), left: Math.min(anchor.col, cell.col),
              bottom: Math.max(anchor.row, cell.row), right: Math.max(anchor.col, cell.col),
            };
    const areas = [...this.state.areas];
    areas[this.state.activeAreaIndex] = extended;
    // Excel keeps the ActiveCell at the drag/Shift origin while only the Area
    // extent follows the pointer. The opposite corner is already encoded by
    // the normalized Area, so moving activeCell here would conflate focus with
    // the range endpoint and paint the focus border on the wrong cell.
    this.state = normalizeSelectionState({ ...this.state, areas });
  }

  snapshot(): XlsxSelectionState | null {
    return this.state ? structuredClone(this.state) : null;
  }

  headerHighlight(): {
    selectedRowRange: { start: number; end: number; strong: boolean } | null;
    selectedColRange: { start: number; end: number; strong: boolean } | null;
  } {
    const area = this.activeArea;
    if (!area) {
      return { selectedRowRange: null, selectedColRange: null };
    }
    const r1 = area.kind === 'cells' ? area.top : area.kind === 'rows' ? area.firstRow : 1;
    const r2 = area.kind === 'cells' ? area.bottom : area.kind === 'rows' ? area.lastRow : Number.MAX_SAFE_INTEGER;
    const c1 = area.kind === 'cells' ? area.left : area.kind === 'columns' ? area.firstColumn : 1;
    const c2 = area.kind === 'cells' ? area.right : area.kind === 'columns' ? area.lastColumn : Number.MAX_SAFE_INTEGER;
    const all = Number.MAX_SAFE_INTEGER;
    switch (area.kind) {
      case 'cells':
        return {
          selectedRowRange: { start: r1, end: r2, strong: false },
          selectedColRange: { start: c1, end: c2, strong: false },
        };
      case 'rows':
        return {
          selectedRowRange: { start: r1, end: r2, strong: true },
          selectedColRange: { start: 1, end: all, strong: false },
        };
      case 'columns':
        return {
          selectedRowRange: { start: 1, end: all, strong: false },
          selectedColRange: { start: c1, end: c2, strong: true },
        };
      case 'sheet':
        return {
          selectedRowRange: { start: 1, end: all, strong: true },
          selectedColRange: { start: 1, end: all, strong: true },
        };
    }
  }
}
