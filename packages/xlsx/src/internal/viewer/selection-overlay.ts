import type { Worksheet } from '../../types.js';
import type { XlsxElementContext, XlsxSelectionState } from '../../selection.js';
import { HEADER_W, HEADER_H, getGridGeometryForWorksheet } from '../../renderer.js';
import { projectXlsxElementContext, type XlsxElementHitViewport } from '../../element-context.js';
import { MAX_WORKSHEET_COL, MAX_WORKSHEET_ROW } from '../grid-geometry.js';
import type { SheetOverlayHost } from '../sheet-surface.js';

/** Default cell-selection accent (Google blue), used when no `selectionColor`
 *  option is supplied. */
const DEFAULT_SELECTION_COLOR = '#1a73e8';

/**
 * Derive the selection rectangle's `border` and `background` CSS from a single
 * accent color: the border is the color verbatim and the fill is the same color
 * at 8% opacity via `color-mix`, so any CSS color string (`#rgb`, `rgb(...)`,
 * named) yields a matching translucent fill without the caller computing an
 * rgba. For the default `#1a73e8` this reproduces the historical
 * `rgba(26,115,232,0.08)` fill.
 */
export function selectionOverlayStyle(color: string): { border: string; background: string } {
  return {
    border: `2px solid ${color}`,
    background: `color-mix(in srgb, ${color} 8%, transparent)`,
  };
}

interface SelectionOverlayRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly top: boolean;
  readonly right: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
}

interface SelectionBoundarySegment {
  readonly axis: 'h' | 'v';
  readonly fixed: number;
  readonly start: number;
  readonly end: number;
}

/**
 * Build the single-Area outline from its visible frozen-pane fragments.
 * Splitting collinear edges at every endpoint emits coincident fragment edges
 * only once. Work is bounded by the visible fragment count, not sheet size.
 */
function selectionBoundaryPath(rects: readonly SelectionOverlayRect[]): string {
  const raw: SelectionBoundarySegment[] = [];
  for (const rect of rects) {
    const x2 = rect.x + rect.width;
    const y2 = rect.y + rect.height;
    if (rect.top) raw.push({ axis: 'h', fixed: rect.y, start: rect.x, end: x2 });
    if (rect.right) raw.push({ axis: 'v', fixed: x2, start: rect.y, end: y2 });
    if (rect.bottom) raw.push({ axis: 'h', fixed: y2, start: rect.x, end: x2 });
    if (rect.left) raw.push({ axis: 'v', fixed: rect.x, start: rect.y, end: y2 });
  }

  const groups = new Map<string, SelectionBoundarySegment[]>();
  for (const segment of raw) {
    const key = `${segment.axis}:${segment.fixed}`;
    const group = groups.get(key);
    if (group) group.push(segment);
    else groups.set(key, [segment]);
  }

  const commands: string[] = [];
  for (const segments of groups.values()) {
    const points = [...new Set(segments.flatMap(({ start, end }) => [start, end]))]
      .sort((a, b) => a - b);
    let runStart: number | null = null;
    let runEnd = 0;
    const flush = () => {
      if (runStart === null || runEnd <= runStart) return;
      const { axis, fixed } = segments[0];
      commands.push(axis === 'h'
        ? `M${runStart} ${fixed}H${runEnd}`
        : `M${fixed} ${runStart}V${runEnd}`);
      runStart = null;
    };
    for (let index = 0; index + 1 < points.length; index++) {
      const start = points[index];
      const end = points[index + 1];
      const covered = segments.some((segment) => segment.start < end && segment.end > start);
      if (covered && runStart !== null && start === runEnd) {
        runEnd = end;
      } else {
        flush();
        if (covered) {
          runStart = start;
          runEnd = end;
        }
      }
    }
    flush();
  }
  return commands.join('');
}

let selectionMaskSequence = 0;

type CellRect = { x: number; y: number; w: number; h: number };

/** Selection model, viewport geometry and follow-up chrome the painter reads. */
export interface SelectionOverlayHost {
  readonly ownerDocument: Document;
  readonly canvasArea: HTMLDivElement;
  readonly overlayHost: SheetOverlayHost;
  worksheet(): Worksheet | null;
  currentSheet(): number;
  selectionState(): XlsxSelectionState | null;
  elementContext(): XlsxElementContext | null;
  elementContextViewport(): XlsxElementHitViewport | null;
  selectionColor(): string | undefined;
  scale(): number;
  isRtl(): boolean;
  cellRect(row: number, col: number): CellRect | null;
  screenX(logicalX: number, width: number): number;
  /** Draw the list-validation arrow for the active cell (selection chrome). */
  drawValidationDropdown(): void;
}

/**
 * Paints the Excel-style selection overlay as DOM/SVG over the grid canvas:
 * the union fill of every selected area split across frozen panes, the
 * single-area boundary, the unshaded ActiveCell cut-out and focus border, or
 * the object-context outline when a drawing is selected.
 */
export class SelectionOverlay {
  constructor(private readonly host: SelectionOverlayHost) {}

  /** Rebuild the selection (or object-context) overlay for the current
   *  viewport, then the list-validation arrow on the active cell. */
  update(): void {
    this.host.overlayHost.clearSelection();
    if (this.host.elementContext()) {
      this.drawElementContextOverlay();
      return;
    }
    const state = this.host.selectionState();
    if (!state) return;
    const cs = this.host.scale();
    const ws = this.host.worksheet();
    if (!ws) return;
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);
    const width = this.host.canvasArea.clientWidth;
    const height = this.host.canvasArea.clientHeight;
    const geometry = getGridGeometryForWorksheet(ws);
    // Match renderViewport's physical freeze materialization. A legal freeze
    // count may cover the full sheet; it must never create million-row overlay
    // geometry when only a handful of bands can reach this viewport.
    const effective = geometry.effectiveFrozenBands({
      scale: cs, width, height, headerWidth: HEADER_W, headerHeight: HEADER_H,
      rows: ws.freezeRows ?? 0, cols: ws.freezeCols ?? 0,
    });
    const axes = geometry.axesAtScale(cs);
    const frozenW = axes.col.offsetOf(effective.cols + 1);
    const frozenH = axes.row.offsetOf(effective.rows + 1);
    const xPanes = effective.cols > 0
      ? [
          { first: 1, last: effective.cols, start: headerW, end: Math.min(width, headerW + frozenW) },
          { first: effective.cols + 1, last: MAX_WORKSHEET_COL, start: Math.min(width, headerW + frozenW), end: width },
        ]
      : [{ first: 1, last: MAX_WORKSHEET_COL, start: headerW, end: width }];
    const yPanes = effective.rows > 0
      ? [
          { first: 1, last: effective.rows, start: headerH, end: Math.min(height, headerH + frozenH) },
          { first: effective.rows + 1, last: MAX_WORKSHEET_ROW, start: Math.min(height, headerH + frozenH), end: height },
        ]
      : [{ first: 1, last: MAX_WORKSHEET_ROW, start: headerH, end: height }];
    const visibleXPanes = xPanes.filter((pane) => pane.end > pane.start);
    const visibleYPanes = yPanes.filter((pane) => pane.end > pane.start);
    const selectionColor = this.host.selectionColor() ?? DEFAULT_SELECTION_COLOR;
    const { background } = selectionOverlayStyle(selectionColor);
    const seenFragments = new Set<string>();
    const fillSubpaths: string[] = [];
    const overlayRects: SelectionOverlayRect[] = [];

    for (const area of state.areas) {
      const bounds = area.kind === 'cells'
        ? { top: area.top, bottom: area.bottom, left: area.left, right: area.right,
            topEdge: true, bottomEdge: true, leftEdge: true, rightEdge: true }
        : area.kind === 'rows'
          ? { top: area.firstRow, bottom: area.lastRow, left: 1, right: MAX_WORKSHEET_COL,
              topEdge: true, bottomEdge: true, leftEdge: false, rightEdge: false }
          : area.kind === 'columns'
            ? { top: 1, bottom: MAX_WORKSHEET_ROW, left: area.firstColumn, right: area.lastColumn,
                topEdge: false, bottomEdge: false, leftEdge: true, rightEdge: true }
            : { top: 1, bottom: MAX_WORKSHEET_ROW, left: 1, right: MAX_WORKSHEET_COL,
                topEdge: false, bottomEdge: false, leftEdge: false, rightEdge: false };

      for (const yp of visibleYPanes) for (const xp of visibleXPanes) {
        const top = Math.max(bounds.top, yp.first);
        const bottom = Math.min(bounds.bottom, yp.last);
        const left = Math.max(bounds.left, xp.first);
        const right = Math.min(bounds.right, xp.last);
        if (top > bottom || left > right) continue;
        const tl = this.host.cellRect(top, left);
        const br = this.host.cellRect(bottom, right);
        if (!tl || !br) continue;
        const rawLeft = tl.x;
        const rawTop = tl.y;
        const rawRight = br.x + br.w;
        const rawBottom = br.y + br.h;
        const x = Math.max(rawLeft, xp.start);
        const y = Math.max(rawTop, yp.start);
        const x2 = Math.min(rawRight, xp.end);
        const y2 = Math.min(rawBottom, yp.end);
        const fragmentW = x2 - x;
        const fragmentH = y2 - y;
        if (fragmentW <= 0 || fragmentH <= 0) continue;

        // Whole rows/columns use the outer visible grid along their unbounded
        // axis. Only outer nonempty panes receive caps; frozen seams and clipped
        // ordinary cell ranges remain open. This is viewer selection chrome.
        const topCap = area.kind === 'columns' && yp === visibleYPanes[0];
        const bottomCap = area.kind === 'columns' && yp === visibleYPanes.at(-1);
        const leftCap = area.kind === 'rows' && xp === visibleXPanes[0];
        const rightCap = area.kind === 'rows' && xp === visibleXPanes.at(-1);
        const topBorder = topCap || (bounds.topEdge && top === bounds.top && rawTop >= yp.start);
        const bottomBorder = bottomCap || (bounds.bottomEdge && bottom === bounds.bottom && rawBottom <= yp.end);
        const leftBorder = leftCap || (bounds.leftEdge && left === bounds.left && rawLeft >= xp.start);
        const rightBorder = rightCap || (bounds.rightEdge && right === bounds.right && rawRight <= xp.end);
        const screenLeft = this.host.screenX(x, fragmentW);
        const physicalLeftBorder = this.host.isRtl() ? rightBorder : leftBorder;
        const physicalRightBorder = this.host.isRtl() ? leftBorder : rightBorder;
        const fragmentKey = [
          screenLeft, y, fragmentW, fragmentH,
          topBorder, physicalRightBorder, bottomBorder, physicalLeftBorder,
        ].join('|');
        if (seenFragments.has(fragmentKey)) continue;
        seenFragments.add(fragmentKey);
        // Paint every fragment as a subpath in one SVG fill operation. With a
        // single non-zero fill, overlapping selection areas form a visual union
        // instead of stacking translucent backgrounds and becoming darker.
        fillSubpaths.push(
          `M${screenLeft} ${y}h${fragmentW}v${fragmentH}h${-fragmentW}Z`,
        );
        // Inset only the new caps by half the centered 2 CSS-pixel stroke so
        // clipping keeps the whole stroke, including the physical RTL left edge.
        // Fill and authored geometry remain unchanged.
        const insetX = Math.min(1, fragmentW / 2);
        const insetY = Math.min(1, fragmentH / 2);
        const outlineLeft = screenLeft + ((this.host.isRtl() ? rightCap : leftCap) ? insetX : 0);
        const outlineRight = screenLeft + fragmentW - ((this.host.isRtl() ? leftCap : rightCap) ? insetX : 0);
        const outlineTop = y + (topCap ? insetY : 0);
        const outlineBottom = y2 - (bottomCap ? insetY : 0);
        overlayRects.push({
          x: outlineLeft,
          y: outlineTop,
          width: outlineRight - outlineLeft,
          height: outlineBottom - outlineTop,
          top: topBorder,
          right: physicalRightBorder,
          bottom: bottomBorder,
          left: physicalLeftBorder,
        });
      }
    }

    if (fillSubpaths.length > 0) {
      const svgNamespace = 'http://www.w3.org/2000/svg';
      const svg = this.host.ownerDocument.createElementNS(svgNamespace, 'svg');
      svg.setAttribute('data-xlsx-selection-fill', '');
      svg.style.cssText =
        'position:absolute;inset:0;width:100%;height:100%;overflow:hidden;pointer-events:none;';
      const isMultipleAreaSelection = state.areas.length > 1;
      const activeRect = this.host.cellRect(state.activeCell.row, state.activeCell.col);
      const maskId = `xlsx-selection-mask-${++selectionMaskSequence}`;
      const defs = this.host.ownerDocument.createElementNS(svgNamespace, 'defs');
      const mask = this.host.ownerDocument.createElementNS(svgNamespace, 'mask');
      mask.setAttribute('id', maskId);
      mask.setAttribute('maskUnits', 'userSpaceOnUse');
      mask.setAttribute('x', '0');
      mask.setAttribute('y', '0');
      mask.setAttribute('width', String(width));
      mask.setAttribute('height', String(height));
      const selectedPath = this.host.ownerDocument.createElementNS(svgNamespace, 'path');
      selectedPath.setAttribute('d', fillSubpaths.join(''));
      selectedPath.setAttribute('fill', '#fff');
      mask.appendChild(selectedPath);

      // Excel leaves ActiveCell unshaded so it remains distinct from the
      // selected cells. ActiveCell stays at the drag origin; only the Area's
      // opposite corner changes during extension.
      if (activeRect) {
        for (const yp of yPanes) for (const xp of xPanes) {
          const clippedX = Math.max(activeRect.x, xp.start);
          const clippedY = Math.max(activeRect.y, yp.start);
          const clippedX2 = Math.min(activeRect.x + activeRect.w, xp.end);
          const clippedY2 = Math.min(activeRect.y + activeRect.h, yp.end);
          if (clippedX2 <= clippedX || clippedY2 <= clippedY) continue;
          const cutout = this.host.ownerDocument.createElementNS(svgNamespace, 'rect');
          cutout.setAttribute('data-xlsx-active-cell-cutout', '');
          cutout.setAttribute('x', String(this.host.screenX(clippedX, clippedX2 - clippedX)));
          cutout.setAttribute('y', String(clippedY));
          cutout.setAttribute('width', String(clippedX2 - clippedX));
          cutout.setAttribute('height', String(clippedY2 - clippedY));
          cutout.setAttribute('fill', '#000');
          mask.appendChild(cutout);
        }
      }
      defs.appendChild(mask);
      svg.appendChild(defs);

      const fill = this.host.ownerDocument.createElementNS(svgNamespace, 'rect');
      fill.setAttribute('x', '0');
      fill.setAttribute('y', '0');
      fill.setAttribute('width', String(width));
      fill.setAttribute('height', String(height));
      fill.setAttribute('fill', background);
      fill.setAttribute('mask', `url(#${maskId})`);
      svg.appendChild(fill);

      const boundaryPath = isMultipleAreaSelection ? '' : selectionBoundaryPath(overlayRects);
      if (boundaryPath) {
        const boundary = this.host.ownerDocument.createElementNS(svgNamespace, 'path');
        boundary.setAttribute('data-xlsx-selection-border', '');
        boundary.setAttribute('d', boundaryPath);
        boundary.setAttribute('fill', 'none');
        boundary.setAttribute('stroke', selectionColor);
        boundary.setAttribute('stroke-width', '2');
        boundary.setAttribute('stroke-linecap', 'square');
        boundary.setAttribute('stroke-linejoin', 'miter');
        svg.appendChild(boundary);
      }
      if (activeRect && isMultipleAreaSelection) {
        for (const yp of yPanes) for (const xp of xPanes) {
          const clippedX = Math.max(activeRect.x, xp.start);
          const clippedY = Math.max(activeRect.y, yp.start);
          const clippedX2 = Math.min(activeRect.x + activeRect.w, xp.end);
          const clippedY2 = Math.min(activeRect.y + activeRect.h, yp.end);
          if (clippedX2 <= clippedX || clippedY2 <= clippedY) continue;
          const focus = this.host.ownerDocument.createElementNS(svgNamespace, 'rect');
          focus.setAttribute('data-xlsx-active-cell-border', '');
          focus.setAttribute('x', String(this.host.screenX(clippedX, clippedX2 - clippedX)));
          focus.setAttribute('y', String(clippedY));
          focus.setAttribute('width', String(clippedX2 - clippedX));
          focus.setAttribute('height', String(clippedY2 - clippedY));
          focus.setAttribute('fill', 'none');
          focus.setAttribute('stroke', selectionColor);
          focus.setAttribute('stroke-width', '1');
          svg.appendChild(focus);
        }
      }
      this.host.overlayHost.appendSelection(svg);
    }

    // List data-validation dropdown arrow (ECMA-376 §18.3.1.33). Excel shows an
    // in-cell dropdown button only while the cell is *selected* and only for
    // `list`-type rules — so it is drawn here (selection overlay) rather than in
    // the canvas renderer. The button itself is non-interactive
    // (pointer-events:none); clicks are hit-tested against its rect in the
    // pointerdown handler, which opens a panel listing the allowed values
    // (display only — picking a value never changes the cell).
    this.host.drawValidationDropdown();
  }

  private drawElementContextOverlay(): void {
    const context = this.host.elementContext();
    const worksheet = this.host.worksheet();
    const viewport = this.host.elementContextViewport();
    if (!context || !worksheet || !viewport || context.sheetIndex !== this.host.currentSheet()) return;
    const projection = projectXlsxElementContext(worksheet, context, viewport);
    if (!projection) return;
    const clip = this.host.ownerDocument.createElement('div');
    clip.setAttribute('data-xlsx-element-context-clip', '');
    clip.style.cssText =
      `position:absolute;left:${projection.clip.x}px;top:${projection.clip.y}px;` +
      `width:${projection.clip.width}px;height:${projection.clip.height}px;` +
      'overflow:hidden;pointer-events:none;';
    const frame = this.host.ownerDocument.createElement('div');
    frame.setAttribute('data-xlsx-element-context-outline', context.elementType);
    const color = this.host.selectionColor() ?? DEFAULT_SELECTION_COLOR;
    frame.style.cssText =
      `position:absolute;left:${projection.rect.x - projection.clip.x}px;` +
      `top:${projection.rect.y - projection.clip.y}px;` +
      `width:${projection.rect.width}px;height:${projection.rect.height}px;` +
      `box-sizing:border-box;border:2px solid ${color};` +
      `background:color-mix(in srgb, ${color} 6%, transparent);` +
      `transform:rotate(${projection.rotation}deg);transform-origin:center;pointer-events:none;`;
    clip.appendChild(frame);
    this.host.overlayHost.appendSelection(clip);
  }
}
