import type { XlsxSelectionContext, CellAddress, XlsxElementContext, XlsxSelectionState } from '../../selection.js';
import { areaContainsCell } from '../../selection.js';
import type { Worksheet } from '../../types.js';
import type { XlsxViewerOptions, XlsxSheetViewerOptions } from '../../viewer.js';
import { zoomStepScale } from '@silurus/ooxml-core';
import {
  HEADER_W,
  HEADER_H,
  pxToColWidth,
  pxToRowHeight,
  getGridGeometryForWorksheet,
} from '../../renderer.js';
import { formatA1 } from '../../a1.js';
import { GridGeometry, MAX_WORKSHEET_COL, MAX_WORKSHEET_ROW } from '../grid-geometry.js';
import type { OutlineAxis } from '../../outline.js';
import type { SelectionController } from '../sheet-viewer-runtime.js';
import type { CanvasSurface } from '../sheet-surface.js';
import { selectionAutoScrollVelocity } from '../../selection-auto-scroll.js';
import type { CommentPopup } from './comment-popup.js';
import type { HyperlinkDispatcher } from './hyperlink-dispatcher.js';
import type { ValidationPanel } from './validation-panel.js';

/** Half-width (CSS px) of the grab zone around a header border for
 *  drag-to-resize (issue #567), and the minimum size a column/row can be
 *  dragged to (logical px) so a collapsed band keeps a grabbable border. */
const RESIZE_GRAB_PX = 4;
const RESIZE_MIN_PX = 5;

// View-only gesture resource policy, not an Excel/OOXML selection limit. The
// current edit/wire format stores one entry per changed band (plus CSS intent
// for columns). Permit a complete 16,384-column worksheet, and give row batches
// the same maximum fan-out: a 1,048,576-row drag would create 64 times that work
// on every pointer event. Union/count before materialization, reject larger
// gestures without mutation, and report the limit through the host error path.
// The host checks cumulative manual overrides too, so repeated gestures cannot
// grow the resize wire indefinitely. Supporting larger row batches efficiently
// requires an interval-based edit/wire representation.
const MAX_RESIZE_BANDS = 16_384;

/**
 * Pure hit predicate for drag-to-resize (issue #567): given a pointer
 * coordinate `pt` (in the header-strip's CSS-px axis — already RTL-un-mirrored
 * by the caller) and the candidate band trailing edges `edges`, return the band
 * index whose edge is within `grabPx` of `pt`, or `null` if none qualifies.
 *
 * `edges` is the candidate list the caller builds — for the band the pointer is
 * over (`hit`) Excel lets you resize the band whose *trailing* border you grab,
 * so the caller passes both `hit - 1` and `hit` (the neighbour-to-the-far-side
 * and the band itself); the first edge within the grab zone wins, in the order
 * given. An edge that sits at or under the header strip (`edge <= headerExtent`,
 * i.e. scrolled behind the frozen corner) is rejected — you can't grab a border
 * hidden under the header. Kept pure (no DOM, no `this`) so the off-by-one
 * geometry — exact-on-edge, within-grab, just-outside, `[hit-1, hit]` neighbour
 * selection, header rejection — is unit-testable. {@link SelectionInput.getResizeTarget}
 * does the DOM/geometry and calls this.
 */
export function resizeHitIndex(
  pt: number,
  edges: { index: number; edge: number }[],
  grabPx: number,
  headerExtent: number,
): number | null {
  for (const { index, edge } of edges) {
    if (edge <= headerExtent) continue; // scrolled behind the header strip
    if (Math.abs(pt - edge) <= grabPx) return index;
  }
  return null;
}

type CellRect = { x: number; y: number; w: number; h: number };

/** Engine state, collaborators and follow-up work the viewport input drives. */
export interface SelectionInputHost {
  readonly surface: CanvasSurface;
  readonly scrollHost: HTMLDivElement;
  readonly canvasArea: HTMLDivElement;
  readonly hostWindow: Window & typeof globalThis;
  /** Whether the mount delegates viewport movement to a native scroll host. */
  readonly nativeScrollbars: boolean;
  readonly selection: SelectionController;
  readonly comments: CommentPopup;
  readonly hyperlinks: HyperlinkDispatcher;
  readonly validation: ValidationPanel;
  options(): XlsxViewerOptions | XlsxSheetViewerOptions;
  worksheet(): Worksheet | null;
  hasWorkbook(): boolean;
  scale(): number;
  isRtl(): boolean;
  isDestroyed(): boolean;
  cellAt(clientX: number, clientY: number): CellAddress | null;
  cellRect(row: number, col: number): CellRect | null;
  screenX(logicalX: number, width: number): number;
  /** Start-anchored logical scroll offsets (see the engine's viewport). */
  scrollLeft(): number;
  scrollTop(): number;
  setScrollLeft(value: number): void;
  setScrollTop(value: number): void;
  scrollCellIntoView(row: number, col: number): void;
  /** Zoom to `scale`, pivoting on `anchor` (canvasArea px) when non-null. */
  zoomAt(anchor: { x: number; y: number } | null, scale: number): void;
  elementContextAt(clientX: number, clientY: number): XlsxElementContext | null;
  setElementContext(context: XlsxElementContext | null): boolean;
  selectionState(): XlsxSelectionState | null;
  setSelection(ref: string): void;
  getSelectionContext(): XlsxSelectionContext | null;
  copySelection(): Promise<unknown>;
  updateSelectionOverlay(): void;
  updateFindOverlay(): void;
  scheduleRender(): void;
  renderCurrentSheet(): Promise<void>;
  emitSelectionChange(): void;
  emitViewportChange(): void;
  hideCommentPopup(): void;
  hideValidationPanel(): void;
  /** `columnCssPx`: logical CSS px captured by a column drag (view-only
   * canonical width that survives MDW changes). Omitted for other edits. */
  recordSizeOverride(axis: OutlineAxis, index: number, columnCssPx?: number): void;
  assertResizeBudget(axis: OutlineAxis, indices: readonly number[], limit: number): void;
  updateSpacerSize(ws: Worksheet): void;
  refitAutoRowsAfterColumnResize(): void;
  reportError(error: unknown): void;
}

/**
 * Pointer, wheel, focus and keyboard input for the sheet viewport: cell,
 * row, column and sheet selection (click, Shift/Ctrl extension and drag with
 * edge auto-scroll), deferred touch/pen taps, header-border drag-to-resize,
 * hyperlink and comment activation, object context selection, the context
 * menu, Ctrl/⌘+wheel zoom and Arrow-key navigation. Every listener is
 * registered on the viewport input element and detached by {@link destroy}.
 */
export class SelectionInput {
  private readonly cleanups: Array<() => void> = [];
  // Deferred selection press: committed on pointerup only if the pointer
  // neither moved beyond the tap threshold nor caused a scroll. Used for
  // touch/pen (swipe-to-scroll must not change the cell) and for mouse
  // presses inside the overlay-scrollbar band (a thumb drag must not select
  // the cell underneath).
  private pendingTap:
    | { x: number; y: number; shiftKey: boolean; additiveKey: boolean; pointerId: number }
    | null = null;
  // IX1 — mouse press bookkeeping for hyperlink activation: the down position and
  // the cell under it. On pointerup, if the pointer did not move beyond the tap
  // slop (a genuine click, not a drag-select), a hyperlink on that cell is
  // dispatched. Touch/pen activate through the pendingTap path instead.
  private pendingClick: { x: number; y: number; pointerId: number; cell: CellAddress } | null = null;
  private pendingElementClick:
    | { x: number; y: number; pointerId: number; context: XlsxElementContext }
    | null = null;
  // In-flight column/row resize drag (issue #567). `originScaled` is the fixed
  // LTR edge the resized band grows from (left edge for a column, top for a row)
  // in canvasArea CSS px; `mdw` is captured once so the live px→model-unit
  // conversion is stable across the drag. A resize is a *view-only* adjustment:
  // it mutates the in-memory worksheet's colWidths/rowHeights, never the file.
  resizeDrag:
    | { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number; pointerId: number;
        indices: readonly number[]; worksheet: Worksheet }
    | null = null;
  /** Last captured drag-selection pointer, retained while edge scrolling runs. */
  private selectionAutoScrollPointer:
    | { clientX: number; clientY: number; pointerId: number }
    | null = null;
  private selectionAutoScrollFrame: number | null = null;
  private selectionAutoScrollLastTime: number | null = null;

  constructor(private readonly host: SelectionInputHost) {}

  private get anchorCell(): CellAddress | null {
    return this.host.selection.anchor;
  }

  private get activeCell(): CellAddress | null {
    return this.host.selection.active;
  }

  private get selectionMode(): SelectionController['mode'] {
    return this.host.selection.mode;
  }

  private get isSelecting(): boolean {
    return this.host.selection.dragging;
  }

  private get selectionPointerId(): number | null {
    return this.host.selection.draggingPointerId;
  }

  private on<K extends keyof HTMLElementEventMap>(
    type: K,
    listener: (event: HTMLElementEventMap[K]) => void,
    options?: AddEventListenerOptions | boolean,
  ): void {
    this.cleanups.push(this.host.surface.on(type, listener, options));
  }

  /** A native scroll cancels deferred presses: the press that started them was
   *  a scrollbar-thumb drag (overlay scrollbars) or a touch swipe. */
  cancelDeferredPress(): void {
    this.pendingTap = null;
    this.pendingElementClick = null;
  }

  /** Navigation/reload installs a different sheet projection: discard deferred
   * clicks and end resize capture without refitting the newly displayed sheet. */
  clearSheetGestures(): void {
    this.pendingTap = null;
    this.pendingClick = null;
    this.pendingElementClick = null;
    if (this.resizeDrag) this.finishResize(this.resizeDrag.pointerId, false);
  }

  /** Claim drag-selection ownership and discard deferred gestures from any
   * other pointer that began before this drag. */
  private beginSelectionDrag(pointerId: number): void {
    if (this.pendingTap?.pointerId !== pointerId) this.pendingTap = null;
    if (this.pendingClick?.pointerId !== pointerId) this.pendingClick = null;
    this.host.selection.beginDrag(pointerId);
  }

  /**
   * Returns what the header area contains at the given client coordinates.
   * Returns null when the point is in the cell grid (not a header).
   */
  getHeaderHit(
    clientX: number,
    clientY: number,
  ): { kind: 'corner' } | { kind: 'row'; row: number } | { kind: 'col'; col: number } | null {
    const ws = this.host.worksheet();
    if (!ws) return null;
    const cs = this.host.scale();
    const rect = this.host.canvasArea.getBoundingClientRect();
    // Same RTL un-mirror as getCellAt: map the screen x back to the logical-LTR
    // layout (row header on the left) before the header math below.
    const lx = this.host.screenX(clientX - rect.left, 0);
    const ly = clientY - rect.top;

    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const inRowHeader = lx < headerW;
    const inColHeader = ly < headerH;
    if (!inRowHeader && !inColHeader) return null;
    if (inRowHeader && inColHeader) return { kind: 'corner' };

    const geometry = getGridGeometryForWorksheet(ws);

    if (inRowHeader) {
      // Determine which row was clicked
      const innerY = ly - headerH;
      if (innerY < 0) return { kind: 'corner' };
      const r = geometry.rowAt(innerY, this.host.scrollTop(), cs);
      return r === null ? null : { kind: 'row', row: r };
    }

    // inColHeader
    const innerX = lx - headerW;
    if (innerX < 0) return { kind: 'corner' };
    const c = geometry.colAt(innerX, this.host.scrollLeft(), cs);
    return c === null ? null : { kind: 'col', col: c };
  }

  /**
   * If the pointer sits on a column/row-header border (within {@link
   * RESIZE_GRAB_PX}), return the resize target: which index to resize and the
   * fixed LTR edge it grows from (in canvasArea CSS px). Excel resizes the band
   * whose *trailing* border you grab — the column to the left of a vertical
   * border, the row above a horizontal one — so both that band and its
   * neighbour-to-the-far-side are checked. Geometry comes straight from {@link
   * getCellRect}, so the grab line always coincides with the drawn border at any
   * scroll offset / zoom / RTL. Returns null off the header borders.
   */
  getResizeTarget(
    clientX: number,
    clientY: number,
  ): { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number } | null {
    const ws = this.host.worksheet();
    if (!ws) return null;
    const cs = this.host.scale();
    const rect = this.host.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x to the logical-LTR space getCellRect draws in (the
    // same transform getHeaderHit uses), so the comparison holds for RTL sheets.
    const ptX = this.host.screenX(clientX - rect.left, 0);
    const ptY = clientY - rect.top;
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const geometry = getGridGeometryForWorksheet(ws);
    const mdw = geometry.maximumDigitWidth;
    const axes = geometry.axesAtScale(cs);

    // Column borders live in the column-header strip, right of the corner.
    if (ptY <= headerH && ptX > headerW) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'col') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      // Zero-width runs share the preceding visible band's trailing edge.
      // Locate that owner by offset instead of walking hidden ordinals.
      const previous = axes.col.indexAt(axes.col.offsetOf(hit.col) - 1).index;
      for (const c of [previous, hit.col]) {
        if (c < 1) continue;
        const r = this.host.cellRect(1, c); // x is independent of the row
        if (!r) continue;
        origins.set(c, r.x);
        edges.push({ index: c, edge: r.x + r.w }); // trailing (right) border
      }
      const index = resizeHitIndex(ptX, edges, RESIZE_GRAB_PX, headerW);
      if (index === null) return null;
      return { kind: 'col', index, originScaled: origins.get(index) as number, mdw };
    }

    // Row borders live in the row-header strip, below the corner.
    if (ptX <= headerW && ptY > headerH) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'row') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      const previous = axes.row.indexAt(axes.row.offsetOf(hit.row) - 1).index;
      for (const rIdx of [previous, hit.row]) {
        if (rIdx < 1) continue;
        const r = this.host.cellRect(rIdx, 1); // y is independent of the column
        if (!r) continue;
        origins.set(rIdx, r.y);
        edges.push({ index: rIdx, edge: r.y + r.h }); // trailing (bottom) border
      }
      const index = resizeHitIndex(ptY, edges, RESIZE_GRAB_PX, headerH);
      if (index === null) return null;
      return { kind: 'row', index, originScaled: origins.get(index) as number, mdw };
    }

    return null;
  }

  /** Snapshot same-axis full-band areas only when the grabbed visible owner
   * belongs to one. API/keyboard-created selections also participate in the
   * next pointer drag; cells, sheet selection and outside boundaries stay single.
   * Union discontiguous/overlapping areas, preserve zero-size hidden bands, and
   * treat frozen bands by logical index. Selection/active cell are unchanged,
   * and later selection changes do not redirect this drag. Not an auto-fit API. */
  private resizeIndices(kind: 'col' | 'row', index: number): number[] {
    const ws = this.host.worksheet();
    const ranges = (this.host.selectionState()?.areas ?? []).flatMap((area) =>
      kind === 'col' && area.kind === 'columns' ? [[area.firstColumn, area.lastColumn]]
        : kind === 'row' && area.kind === 'rows' ? [[area.firstRow, area.lastRow]] : []);
    if (!ws || !ranges.some(([first, last]) => index >= first && index <= last)) return [index];
    ranges.sort((a, b) => a[0] - b[0]);
    const union: number[][] = [];
    for (const range of ranges) {
      const previous = union.at(-1);
      if (previous && range[0] <= previous[1] + 1) previous[1] = Math.max(previous[1], range[1]);
      else union.push([...range]);
    }
    const count = union.reduce((total, [first, last]) => total + last - first + 1, 0);
    if (count > MAX_RESIZE_BANDS) {
      throw new RangeError(`A resize gesture may affect at most ${MAX_RESIZE_BANDS} selected bands.`);
    }
    const geometry = getGridGeometryForWorksheet(ws);
    // Geometry also resolves range-encoded column widths and zero defaults;
    // checking only the explicit size dictionary would unhide those bands.
    const axis = kind === 'col' ? geometry.col : geometry.row;
    const indices: number[] = [];
    for (const [first, last] of union) {
      for (let band = first; band <= last; band++) {
        if (axis.sizeOf(band) > 0) indices.push(band);
      }
    }
    return indices;
  }

  /** End input ownership before release/refit, which can fail. Like the existing
   * single-band live resize, pointercancel retains the last completed live size;
   * it does not undo an edit. No subsequent move from that pointer can resize.
   * Render failures happen after the whole batch's model/wire edit is recorded:
   * they report an error and may leave the prior bitmap, never a partial batch. */
  private finishResize(pointerId: number, refit = true): void {
    const drag = this.resizeDrag;
    if (!drag || drag.pointerId !== pointerId) return;
    this.resizeDrag = null;
    try {
      // The UA may already have released capture for pointercancel/lost capture.
      if (this.host.scrollHost.hasPointerCapture?.(pointerId) !== false) {
        this.host.scrollHost.releasePointerCapture(pointerId);
      }
    } catch (error) {
      this.host.reportError(error);
    }
    try {
      if (refit && drag.kind === 'col' && drag.worksheet === this.host.worksheet()) {
        this.host.refitAutoRowsAfterColumnResize();
      }
    } catch (error) {
      this.host.reportError(error);
    }
  }

  /**
   * Apply a live resize drag: size the band from its fixed origin edge to the
   * current pointer, clamp to {@link RESIZE_MIN_PX}, and write the result back
   * into the in-memory worksheet model in its native unit (Excel column widths /
   * points). This is a *view-only* mutation — the file is never written. The
   * memoized axis cache for this sheet is invalidated so every geometry read
   * (spacer, hit-test, overlay, renderer) sees the new size on the next frame.
   */
  applyResize(clientX: number, clientY: number): void {
    const drag = this.resizeDrag;
    const ws = this.host.worksheet();
    if (!drag) return;
    // A delayed move must never edit a replacement projection after navigation
    // or reload. Normally clearSheetGestures releases ownership at installation.
    if (!ws || drag.worksheet !== ws) {
      this.finishResize(drag.pointerId, false);
      return;
    }
    const cs = this.host.scale();
    const rect = this.host.canvasArea.getBoundingClientRect();

    if (drag.kind === 'col') {
      const ptX = this.host.screenX(clientX - rect.left, 0);
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptX - drag.originScaled) / cs));
      const width = pxToColWidth(sizePx, drag.mdw);
      for (const index of drag.indices) {
        ws.colWidths[index] = width;
        // Every target retains the same drag CSS intent through MDW rebind,
        // auto-height clones and worker projections; authored widths stay raw.
        this.host.recordSizeOverride('col', index, sizePx);
      }
    } else {
      const ptY = clientY - rect.top;
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptY - drag.originScaled) / cs));
      const height = pxToRowHeight(sizePx);
      for (const index of drag.indices) {
        ws.rowHeights[index] = height;
        this.host.recordSizeOverride('row', index);
      }
    }

    GridGeometry.invalidate(ws); // sizes changed → rebuild the cumulative-offset axes
    this.host.updateSpacerSize(ws);
    this.host.updateSelectionOverlay();
    // Live resize drag fires per pointermove; coalesce the canvas repaint into
    // one frame. The spacer (scrollbar extent) and overlay updates are cheap DOM
    // writes that must track the drag immediately, so they stay synchronous.
    this.host.scheduleRender();
  }

  private applyPointerSelection(
    clientX: number,
    clientY: number,
    shiftKey: boolean,
    additiveKey: boolean,
    pointerId: number,
    allowDrag: boolean,
  ): void {
    const headerHit = this.getHeaderHit(clientX, clientY);

    if (headerHit) {
      if (headerHit.kind === 'corner') {
        // Select all — no drag extension needed
        this.host.selection.select({ row: 1, col: 1 }, 'all');
        this.host.selection.endDrag();
      } else if (headerHit.kind === 'row') {
        if (shiftKey && this.anchorCell && this.selectionMode === 'rows') {
          this.host.selection.extend({ row: headerHit.row, col: 1 });
        } else {
          const selected = additiveKey
            ? this.host.selection.add({ row: headerHit.row, col: 1 }, 'rows')
            : (this.host.selection.select({ row: headerHit.row, col: 1 }, 'rows'), true);
          if (allowDrag && selected) {
            this.beginSelectionDrag(pointerId);
            this.host.scrollHost.setPointerCapture(pointerId);
          }
        }
      } else {
        if (shiftKey && this.anchorCell && this.selectionMode === 'cols') {
          this.host.selection.extend({ row: 1, col: headerHit.col });
        } else {
          const selected = additiveKey
            ? this.host.selection.add({ row: 1, col: headerHit.col }, 'cols')
            : (this.host.selection.select({ row: 1, col: headerHit.col }, 'cols'), true);
          if (allowDrag && selected) {
            this.beginSelectionDrag(pointerId);
            this.host.scrollHost.setPointerCapture(pointerId);
          }
        }
      }
      this.host.updateSelectionOverlay();
      void this.host.renderCurrentSheet().catch((error) => this.host.reportError(error));
      this.host.emitSelectionChange();
      return;
    }

    const cell = this.host.cellAt(clientX, clientY);
    if (!cell) return;

    let selected = true;
    if (shiftKey && this.anchorCell && this.selectionMode === 'cells') {
      this.host.selection.extend(cell);
    } else {
      selected = additiveKey
        ? this.host.selection.add(cell, 'cells')
        : (this.host.selection.select(cell, 'cells'), true);
    }
    if (allowDrag && selected) {
      this.beginSelectionDrag(pointerId);
      this.host.scrollHost.setPointerCapture(pointerId);
    }
    this.host.updateSelectionOverlay();
    if (this.host.hasWorkbook()) {
      this.host.renderCurrentSheet().catch((error) => this.host.reportError(error));
    }
    this.host.emitSelectionChange();
  }

  /** Browser-visible input box, excluding classic native scrollbar gutters. */
  private viewportInputBounds(): { left: number; top: number; width: number; height: number } {
    const rect = this.host.canvasArea.getBoundingClientRect();
    const left = rect.left + this.host.scrollHost.clientLeft;
    const top = rect.top + this.host.scrollHost.clientTop;
    const availableWidth = Math.max(0, rect.width - this.host.scrollHost.clientLeft);
    const availableHeight = Math.max(0, rect.height - this.host.scrollHost.clientTop);
    return {
      left,
      top,
      width: Math.min(availableWidth, this.host.scrollHost.clientWidth || availableWidth),
      height: Math.min(availableHeight, this.host.scrollHost.clientHeight || availableHeight),
    };
  }

  /** Extend the active drag selection to the pointer's cell. Captured pointers
   * outside the canvas and auto-scroll ticks clamp to the visible data edge so
   * selection never jumps ahead of the viewport. */
  private extendDragSelection(
    clientX: number,
    clientY: number,
    clampToViewport: boolean,
  ): boolean {
    let pointerX = clientX;
    let pointerY = clientY;
    const bounds = this.viewportInputBounds();
    const outsideViewport = clientX < bounds.left || clientX >= bounds.left + bounds.width ||
      clientY < bounds.top || clientY >= bounds.top + bounds.height;
    if (clampToViewport || outsideViewport) {
      const cs = this.host.scale();
      const headerW = Math.round(HEADER_W * cs);
      const headerH = Math.round(HEADER_H * cs);
      const dataLeft = bounds.left + (this.host.isRtl() ? 0 : headerW);
      const dataRight = bounds.left + bounds.width - (this.host.isRtl() ? headerW : 0);
      pointerX = Math.min(dataRight - 1, Math.max(dataLeft + 1, pointerX));
      pointerY = Math.min(
        bounds.top + bounds.height - 1,
        Math.max(bounds.top + headerH + 1, pointerY),
      );
    }

    if (this.selectionMode === 'rows') {
      const hit = clampToViewport ? null : this.getHeaderHit(pointerX, pointerY);
      const row = hit?.kind === 'row'
        ? hit.row
        : this.host.cellAt(pointerX, pointerY)?.row;
      if (!row || row === this.activeCell?.row) return false;
      this.host.selection.extend({ row, col: 1 });
      return true;
    }

    if (this.selectionMode === 'cols') {
      const hit = clampToViewport ? null : this.getHeaderHit(pointerX, pointerY);
      const col = hit?.kind === 'col'
        ? hit.col
        : this.host.cellAt(pointerX, pointerY)?.col;
      if (!col || col === this.activeCell?.col) return false;
      this.host.selection.extend({ row: 1, col });
      return true;
    }

    const cell = this.host.cellAt(pointerX, pointerY);
    if (!cell || (cell.row === this.activeCell?.row && cell.col === this.activeCell?.col)) {
      return false;
    }
    this.host.selection.extend(cell);
    return true;
  }

  private selectionAutoScrollSpeed(): { x: number; y: number } {
    const pointer = this.selectionAutoScrollPointer;
    if (!pointer) return { x: 0, y: 0 };
    const bounds = this.viewportInputBounds();
    return selectionAutoScrollVelocity(
      { x: pointer.clientX - bounds.left, y: pointer.clientY - bounds.top },
      { width: bounds.width, height: bounds.height },
      this.host.isRtl(),
      this.selectionMode,
    );
  }

  private trackSelectionAutoScroll(e: PointerEvent): void {
    if (e.pointerId !== this.selectionPointerId) return;
    this.selectionAutoScrollPointer = {
      clientX: e.clientX,
      clientY: e.clientY,
      pointerId: e.pointerId,
    };
    const speed = this.selectionAutoScrollSpeed();
    if (speed.x === 0 && speed.y === 0) {
      this.stopSelectionAutoScroll();
      return;
    }
    if (this.selectionAutoScrollFrame !== null) return;
    this.selectionAutoScrollLastTime = null;
    this.selectionAutoScrollFrame = this.host.hostWindow.requestAnimationFrame(
      (time) => this.runSelectionAutoScroll(time),
    );
  }

  private runSelectionAutoScroll(time: number): void {
    this.selectionAutoScrollFrame = null;
    const pointer = this.selectionAutoScrollPointer;
    if (
      !pointer ||
      pointer.pointerId !== this.selectionPointerId ||
      !this.isSelecting ||
      this.host.isDestroyed()
    ) {
      this.stopSelectionAutoScroll();
      return;
    }

    const speed = this.selectionAutoScrollSpeed();
    if (speed.x === 0 && speed.y === 0) {
      this.stopSelectionAutoScroll();
      return;
    }

    const previousTime = this.selectionAutoScrollLastTime;
    const elapsedSeconds = previousTime === null
      ? 1 / 60
      : Math.min(0.05, Math.max(0, time - previousTime) / 1000);
    this.selectionAutoScrollLastTime = time;

    const beforeX = this.host.scrollLeft();
    const beforeY = this.host.scrollTop();
    this.host.setScrollLeft(beforeX + speed.x * elapsedSeconds);
    this.host.setScrollTop(beforeY + speed.y * elapsedSeconds);
    const moved = this.host.scrollLeft() !== beforeX || this.host.scrollTop() !== beforeY;
    const extended = moved && this.extendDragSelection(pointer.clientX, pointer.clientY, true);

    if (moved) {
      this.host.updateSelectionOverlay();
      this.host.updateFindOverlay();
      this.host.scheduleRender();
      this.host.emitViewportChange();
      if (extended) this.host.emitSelectionChange();
    }

    if (!moved) {
      this.stopSelectionAutoScroll();
      return;
    }
    this.selectionAutoScrollFrame = this.host.hostWindow.requestAnimationFrame(
      (nextTime) => this.runSelectionAutoScroll(nextTime),
    );
  }

  private stopSelectionAutoScroll(): void {
    if (this.selectionAutoScrollFrame !== null) {
      this.host.hostWindow.cancelAnimationFrame(this.selectionAutoScrollFrame);
      this.selectionAutoScrollFrame = null;
    }
    this.selectionAutoScrollPointer = null;
    this.selectionAutoScrollLastTime = null;
  }

  private contextMenuTargetIsSelected(clientX: number, clientY: number): boolean {
    const selection = this.host.selectionState();
    if (!selection) return false;
    const header = this.getHeaderHit(clientX, clientY);
    if (header?.kind === 'corner') {
      return selection.areas.some((area) => area.kind === 'sheet');
    }
    if (header?.kind === 'row') {
      return selection.areas.some((area) => area.kind === 'sheet' ||
        (area.kind === 'rows' && header.row >= area.firstRow && header.row <= area.lastRow));
    }
    if (header?.kind === 'col') {
      return selection.areas.some((area) => area.kind === 'sheet' ||
        (area.kind === 'columns' &&
          header.col >= area.firstColumn && header.col <= area.lastColumn));
    }
    const cell = this.host.cellAt(clientX, clientY);
    return cell !== null && selection.areas.some((area) => areaContainsCell(area, cell));
  }

  private resolveContextMenuContext(event: MouseEvent): Promise<XlsxSelectionContext | null> {
    if (this.host.isDestroyed()) return Promise.resolve(null);
    const element = this.host.elementContextAt(event.clientX, event.clientY);
    if (element) {
      this.host.setElementContext(element);
    } else {
      this.host.setElementContext(null);
      if (!this.contextMenuTargetIsSelected(event.clientX, event.clientY)) {
        this.applyPointerSelection(event.clientX, event.clientY, false, false, -1, false);
      }
    }
    const context = this.host.getSelectionContext();
    return Promise.resolve(context ? structuredClone(context) : null);
  }

  /** Install the viewport's pointer, wheel, focus and keyboard listeners. */
  install(): void {
    // Distance (CSS px) beyond which a touch/pen pointerdown→pointerup is treated as a swipe (scroll), not a tap.
    const TAP_SLOP = 8;

    if (this.host.options().onContextMenu) {
      this.on('contextmenu', (event: MouseEvent) => {
        let context: Promise<XlsxSelectionContext | null> | undefined;
        this.host.options().onContextMenu?.({
          originalEvent: event,
          getContext: () => context ??= this.resolveContextMenuContext(event),
        });
      });
    }

    this.on('pointerdown', (e: PointerEvent) => {
      this.host.scrollHost.focus?.({ preventScroll: true });
      if (e.button !== 0) return;
      if (this.resizeDrag) return; // another pointer cannot steal a live resize
      if (this.isSelecting && e.pointerId !== this.selectionPointerId) return;

      // Drag-to-resize a column/row from its header border (issue #567). Checked
      // before selection so grabbing the border never moves the cell selection.
      // Gated by the `resizable` option (default true); when off, a header-border
      // press falls through to normal selection behavior.
      const resize = (this.host.options().resizable ?? true)
        ? this.getResizeTarget(e.clientX, e.clientY)
        : null;
      if (resize) {
        e.preventDefault();
        try {
          const worksheet = this.host.worksheet()!;
          const indices = this.resizeIndices(resize.kind, resize.index);
          this.host.assertResizeBudget(resize.kind, indices, MAX_RESIZE_BANDS);
          // A failed capture must not leave a partially started resize alive.
          this.host.scrollHost.setPointerCapture(e.pointerId);
          this.resizeDrag = { ...resize, indices, worksheet, pointerId: e.pointerId };
        } catch (error) {
          this.host.reportError(error);
          return;
        }
        this.host.hideCommentPopup();
        return;
      }

      // List-validation dropdown arrow: if the press lands on the (display-only)
      // arrow button drawn on the active cell, toggle the value panel instead of
      // re-selecting the cell. The arrow's rect is in canvasArea space, so map
      // the client point through canvasArea's box.
      if (this.host.validation.hitsArrow(e.clientX, e.clientY)) {
        e.preventDefault();
        this.host.validation.toggle();
        return;
      }

      // A pointerdown on the native scrollbar must not move the cell
      // selection — dragging the thumb would otherwise select whatever cell
      // sits underneath it. Two scrollbar styles need different handling:
      // classic scrollbars reserve layout space, so the press lands in the
      // band between the content box (clientWidth/Height) and the border-box
      // edge and can be rejected exactly; OS overlay scrollbars (macOS
      // "show when scrolling") float over the content without affecting
      // client sizes, so a press near a scrollable edge is geometrically
      // indistinguishable from a cell click. For that case we defer the
      // selection to pointerup via the pendingTap path and cancel it when a
      // scroll event arrives first (the press was a thumb drag). A plain
      // click in the band still selects the cell on release.
      const hostRect = this.host.scrollHost.getBoundingClientRect();
      const localX = e.clientX - hostRect.left - this.host.scrollHost.clientLeft;
      const localY = e.clientY - hostRect.top - this.host.scrollHost.clientTop;
      if (localX >= this.host.scrollHost.clientWidth || localY >= this.host.scrollHost.clientHeight) {
        return; // classic scrollbar gutter
      }
      // Overlay scrollbar hit band (~15 CSS px on macOS / Windows 11).
      const OVERLAY_SCROLLBAR_BAND = 16;
      const inOverlayBand = this.host.nativeScrollbars && (
        (this.host.scrollHost.scrollWidth > this.host.scrollHost.clientWidth &&
          this.host.scrollHost.clientHeight - localY <= OVERLAY_SCROLLBAR_BAND) ||
        (this.host.scrollHost.scrollHeight > this.host.scrollHost.clientHeight &&
          this.host.scrollHost.clientWidth - localX <= OVERLAY_SCROLLBAR_BAND));

      const elementContext = this.host.elementContextAt(e.clientX, e.clientY);
      if (elementContext) {
        this.pendingTap = null;
        this.pendingClick = null;
        this.pendingElementClick = {
          x: e.clientX,
          y: e.clientY,
          pointerId: e.pointerId,
          context: elementContext,
        };
        return;
      }
      // A cell/header/empty-space press leaves object focus and returns the
      // authoritative context to the existing cell-selection state.
      this.host.setElementContext(null);

      // Touch / pen: defer selection until pointerup so swipe-to-scroll doesn't change the cell.
      // Mouse: select immediately to preserve drag-to-extend behavior.
      if (e.pointerType !== 'mouse' || inOverlayBand) {
        this.pendingTap = {
          x: e.clientX,
          y: e.clientY,
          shiftKey: e.shiftKey,
          additiveKey: e.ctrlKey || e.metaKey,
          pointerId: e.pointerId,
        };
        return;
      }

      // IX1 — remember the cell under a mouse press so a click (no drag) can
      // activate its hyperlink on release. Recorded before selection so a
      // shift-click extend still tracks the destination cell.
      const downCell = this.host.cellAt(e.clientX, e.clientY);
      this.pendingClick = downCell
        ? { x: e.clientX, y: e.clientY, pointerId: e.pointerId, cell: downCell }
        : null;

      this.applyPointerSelection(
        e.clientX,
        e.clientY,
        e.shiftKey,
        e.ctrlKey || e.metaKey,
        e.pointerId,
        true,
      );
    });

    this.on('pointermove', (e: PointerEvent) => {
      // Live column/row resize takes priority over every other pointer behavior.
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        e.preventDefault();
        this.applyResize(e.clientX, e.clientY);
        return;
      }

      // Resize-handle affordance: show the col/row-resize cursor when hovering a
      // header border (mouse only — touch/pen have no hover). Skipped mid-select
      // and when the `resizable` option (default true) is off, so no resize
      // cursor is shown when drag-resize is disabled.
      if (e.pointerType === 'mouse' && !this.isSelecting && (this.host.options().resizable ?? true)) {
        const rt = this.getResizeTarget(e.clientX, e.clientY);
        this.host.scrollHost.style.cursor = rt ? (rt.kind === 'col' ? 'col-resize' : 'row-resize') : '';
        if (rt) {
          this.host.hideCommentPopup();
          return;
        }
      }

      // Cancel a pending tap once the pointer moves beyond the slop — the user is scrolling.
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingTap = null;
        }
      }

      // IX1 — a mouse press that turns into a drag (beyond the slop) is a
      // selection, not a hyperlink click: drop the pending activation.
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingClick.x;
        const dy = e.clientY - this.pendingClick.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingClick = null;
        }
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingElementClick.x;
        const dy = e.clientY - this.pendingElementClick.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) this.pendingElementClick = null;
      }

      // Comment hover popup (mouse only — touch/pen have no hover, so they get
      // the popup on selection instead, below). Suppressed while drag-selecting
      // so the popup doesn't fight the selection rect. A header hover hides it.
      if (e.pointerType === 'mouse' && !this.isSelecting) {
        const hovered = this.host.cellAt(e.clientX, e.clientY);
        if (hovered) this.host.comments.scheduleForCell(hovered);
        else this.host.hideCommentPopup();
        // IX1 — pointer cursor over a hyperlinked cell. Reached only when the
        // pointer is NOT over a resize border (that path returns above), so the
        // resize cursor is never clobbered. Otherwise clear back to default.
        this.host.scrollHost.style.cursor =
          hovered && this.host.hyperlinks.at(hovered) ? 'pointer' : '';
      }

      if (!this.isSelecting || e.pointerId !== this.selectionPointerId) return;

      this.trackSelectionAutoScroll(e);
      if (!this.extendDragSelection(e.clientX, e.clientY, false)) return;

      this.host.updateSelectionOverlay();
      // Drag-select fires per pointermove; coalesce the canvas repaint (the
      // header-highlight bands the renderer draws) into one frame. The overlay
      // rect and the selection-change callback stay synchronous.
      this.host.scheduleRender();
      this.host.emitSelectionChange();
    });

    this.on('pointerup', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        this.finishResize(e.pointerId);
        return;
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        const pending = this.pendingElementClick;
        this.pendingElementClick = null;
        const dx = e.clientX - pending.x;
        const dy = e.clientY - pending.y;
        const current = dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP
          ? this.host.elementContextAt(e.clientX, e.clientY)
          : null;
        if (
          current &&
          current.sheetIndex === pending.context.sheetIndex &&
          current.elementType === pending.context.elementType &&
          current.elementIndex === pending.context.elementIndex &&
          current.shapeIndex === pending.context.shapeIndex
        ) this.host.setElementContext(current);
        return;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP) {
          this.applyPointerSelection(
            e.clientX,
            e.clientY,
            this.pendingTap.shiftKey,
            this.pendingTap.additiveKey,
            e.pointerId,
            false,
          );
          // Touch / pen have no hover, so surface the comment popup on a tap
          // (the active cell after the selection commit). Mouse uses hover.
          if (e.pointerType !== 'mouse' && this.activeCell) {
            const comment = this.host.comments.commentAt(this.activeCell);
            if (comment) {
              this.host.hideCommentPopup();
              void this.host.comments.show(this.activeCell, comment)
                .catch((error) => this.host.reportError(error));
            } else {
              this.host.hideCommentPopup();
            }
          }
          // IX1 — a touch/pen tap on a hyperlinked cell activates it.
          if (this.activeCell) this.host.hyperlinks.dispatch(this.activeCell);
        }
        this.pendingTap = null;
      }
      const endsSelectionDrag = e.pointerId === this.selectionPointerId;
      if (endsSelectionDrag) this.stopSelectionAutoScroll();
      // IX1 — a mouse click (press+release without a drag) on a hyperlinked cell
      // activates it. The release must still land on the same cell the press did.
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingClick.x;
        const dy = e.clientY - this.pendingClick.y;
        const upCell = this.host.cellAt(e.clientX, e.clientY);
        if (
          dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP &&
          upCell &&
          upCell.row === this.pendingClick.cell.row &&
          upCell.col === this.pendingClick.cell.col
        ) {
          this.host.hyperlinks.dispatch(this.pendingClick.cell);
        }
        this.pendingClick = null;
      }
      if (endsSelectionDrag) this.host.selection.endDrag(e.pointerId);
    });

    this.on('pointercancel', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        this.finishResize(e.pointerId);
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        this.pendingTap = null;
      }
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        this.pendingClick = null;
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        this.pendingElementClick = null;
      }
      if (e.pointerId === this.selectionPointerId) {
        this.stopSelectionAutoScroll();
        this.host.selection.endDrag(e.pointerId);
      }
    });

    // Ctrl/⌘ + mouse wheel (and trackpad pinch, which the browser reports as a
    // ctrl-wheel) zooms the grid, matching Excel. preventDefault stops the
    // browser's own page zoom. A plain wheel still scrolls the grid natively.
    // The step is exponential in mode-normalized wheel distance (see
    // zoomStepScale), so a trackpad pinch — a high-frequency stream of
    // small-deltaY events — does not zoom away; the total zoom tracks the gesture
    // distance, not the event count, while a mouse wheel remains a gentle 10%.
    this.on(
      'wheel',
      (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) {
          if (!this.host.nativeScrollbars) {
            e.preventDefault();
            const unit = e.deltaMode === WheelEvent.DOM_DELTA_LINE
              ? 16
              : e.deltaMode === WheelEvent.DOM_DELTA_PAGE
                ? Math.max(1, this.host.scrollHost.clientHeight)
                : 1;
            const horizontal = (e.shiftKey ? e.deltaY : e.deltaX) * unit;
            const vertical = (e.shiftKey ? 0 : e.deltaY) * unit;
            this.host.setScrollLeft(this.host.scrollLeft() + horizontal);
            this.host.setScrollTop(this.host.scrollTop() + vertical);
            this.host.scheduleRender();
            this.host.updateSelectionOverlay();
            this.host.updateFindOverlay();
            this.host.emitViewportChange();
          }
          return;
        }
        e.preventDefault();
        if (e.deltaY === 0) return;
        // Pointer-anchored zoom: pivot on the cursor, not the top-left corner.
        // Record the pointer relative to the grid's top-left (canvasArea rect,
        // which the scrollHost overlays with inset:0) so `setScale` keeps the
        // cell under the cursor fixed. `scrollHost` and `canvasArea` share a rect.
        // A malformed event (no clientX/Y) yields a non-finite anchor; drop it so
        // `setScale` falls back to the historical START-anchored preservation.
        const { x: ax, y: ay } = this.host.surface.localPoint(e.clientX, e.clientY);
        this.host.zoomAt(
          Number.isFinite(ax) && Number.isFinite(ay) ? { x: ax, y: ay } : null,
          zoomStepScale(this.host.scale(), e.deltaY, e.deltaMode),
        );
      },
      { passive: false },
    );

    this.on('pointerleave', (event: PointerEvent) => {
      const next = event.relatedTarget as Node | null;
      if (next && this.host.comments.contains(next)) return;
      this.host.hideCommentPopup();
    });

    // A canvas-backed sheet has no native focused cell. Establish the ordinary
    // A1 selection when its viewport receives keyboard focus, then reuse the
    // public selection contract for Arrow-key movement below.
    this.on('focus', () => {
      if (this.host.worksheet() && !this.activeCell) this.host.setSelection('A1');
    });

    this.on('keydown', (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        if (e.defaultPrevented || e.isComposing) return;
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (target?.isContentEditable || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
        e.preventDefault();
        void this.host.copySelection();
      } else if (
        !e.defaultPrevented && !e.isComposing &&
        !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey &&
        (e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
          e.key === 'ArrowLeft' || e.key === 'ArrowRight')
      ) {
        const current = this.activeCell;
        const rowDelta = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;
        const colDelta = e.key === 'ArrowLeft'
          ? (this.host.isRtl() ? 1 : -1)
          : e.key === 'ArrowRight'
            ? (this.host.isRtl() ? -1 : 1)
            : 0;
        const next = current ? {
          row: Math.max(1, Math.min(MAX_WORKSHEET_ROW, current.row + rowDelta)),
          col: Math.max(1, Math.min(MAX_WORKSHEET_COL, current.col + colDelta)),
        } : { row: 1, col: 1 };
        e.preventDefault();
        this.host.hideCommentPopup();
        const ref = formatA1(next.row, next.col);
        this.host.setSelection(ref);
        // Selection already schedules the paint. Reuse the ordinary viewport
        // geometry without starting a second immediate render for every key.
        this.host.scrollCellIntoView(next.row, next.col);
        this.host.updateSelectionOverlay();
        this.host.updateFindOverlay();
        this.host.emitViewportChange();
      } else if (e.key === 'Escape' && this.host.validation.isOpen()) {
        this.host.hideValidationPanel();
      } else if (e.key === 'Escape' && this.host.comments.isOpen()) {
        this.host.hideCommentPopup();
      } else if (
        e.key === 'Enter' && this.activeCell &&
        !e.defaultPrevented && !e.isComposing &&
        !e.ctrlKey && !e.metaKey && !e.altKey
      ) {
        const comment = this.host.comments.commentAt(this.activeCell);
        if (comment) {
          e.preventDefault();
          this.host.hideCommentPopup();
          void this.host.comments.show(this.activeCell, comment)
            .catch((error) => this.host.reportError(error));
        }
      }
    });
  }

  /** Teardown: stop edge scrolling, drop pending gestures and detach every
   *  viewport listener this input installed. */
  destroy(): void {
    this.stopSelectionAutoScroll();
    this.pendingTap = null;
    this.pendingClick = null;
    this.pendingElementClick = null;
    this.resizeDrag = null;
    for (const cleanup of this.cleanups.splice(0)) cleanup();
  }
}
