import type { FindHighlightColors, FindMatch, FindMatchesOptions } from '@silurus/ooxml-core';
import type { Worksheet } from '../../types.js';
import type { XlsxWorkbook } from '../../workbook.js';
import { acquireXlsxWorksheet, xlsxCellDisplayTexts } from '../../workbook.js';
import { HEADER_W, HEADER_H, getGridGeometryForWorksheet } from '../../renderer.js';
import { XlsxFindController, type FindCell, type XlsxMatchLocation } from '../../find.js';
import type { SheetOverlayHost } from '../sheet-surface.js';

const DEFAULT_FIND_HIGHLIGHT = 'color-mix(in srgb, #ffb300 8%, transparent)';
const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'color-mix(in srgb, #fb8c00 8%, transparent)';

/** Resolve an XLSX find box without altering a caller-provided CSS background. */
export function findHighlightOverlayStyle(
  active: boolean,
  colors: FindHighlightColors = {},
): { border: string; background: string } {
  const accent = active ? '#fb8c00' : '#ffb300';
  const custom = active ? colors.active : colors.match;
  const background = custom ?? (active ? DEFAULT_FIND_ACTIVE_HIGHLIGHT : DEFAULT_FIND_HIGHLIGHT);
  return { border: `2px solid ${custom ?? accent}`, background };
}

type CellRect = { x: number; y: number; w: number; h: number };

/** Workbook access, geometry and navigation the find adapter uses. */
export interface FindAdapterHost {
  readonly ownerDocument: Document;
  readonly overlayHost: SheetOverlayHost;
  workbook(): XlsxWorkbook | null;
  sheetCount(): number;
  worksheet(): Worksheet | null;
  currentSheet(): number;
  scale(): number;
  highlightColors(): FindHighlightColors | undefined;
  cellRect(row: number, col: number): CellRect | null;
  screenX(logicalX: number, width: number): number;
  goToSheet(index: number): Promise<void>;
  /** Scroll so a cell is in view (Excel's find behaviour: in-view cells stay). */
  scrollCellIntoView(row: number, col: number): void;
}

/**
 * IX2 find for the workbook viewer: adapts the format-neutral
 * {@link XlsxFindController} to worksheet leases (the whole-sheet search
 * source) and draws the find-highlight overlay for the displayed sheet.
 */
export class FindAdapter {
  /** Find state (matches + active cursor); survives sheet switches. */
  readonly controller: XlsxFindController;

  constructor(private readonly host: FindAdapterHost) {
    this.controller = new XlsxFindController(
      () => host.sheetCount(),
      (sheet) => host.workbook()?.sheetNames[sheet] ?? '',
      (sheet) => this.collectSheetCells(sheet),
    );
  }

  /** Every non-empty cell of a sheet with its rendered display text (the find
   *  source). Reads the parsed worksheet model directly — no render — so search
   *  covers the whole sheet, not just the on-screen viewport. */
  private async collectSheetCells(sheet: number): Promise<FindCell[]> {
    const wb = this.host.workbook();
    if (!wb) return [];
    const lease = await acquireXlsxWorksheet(wb, sheet);
    try {
      const ws = lease.worksheet;
      const cells: FindCell[] = [];
      for (const row of ws.rows) {
        for (const cell of row.cells) {
          // The unconstrained display text, then each shorter General form
          // paint may show at the cell's width (issue #2179).
          const [text = '', ...alternates] = wb[xlsxCellDisplayTexts](ws, cell);
          if (text === '') continue;
          cells.push(alternates.length > 0
            ? { row: cell.row, col: cell.col, text, alternates }
            : { row: cell.row, col: cell.col, text });
        }
      }
      return cells;
    } finally {
      lease.release();
    }
  }

  /** Drop matches and cursor (new workbook, clearFind, teardown). */
  invalidate(): void {
    this.controller.invalidate();
  }

  /**
   * Redraw the find-highlight overlay: one translucent box per matched cell on
   * the current sheet, the active match in a stronger colour. Uses the SAME
   * cell rect + screenX + header/frozen clamp the selection overlay uses, so a
   * box lands exactly on the drawn cell at any scroll offset / zoom / RTL.
   * Rebuilt on every render and scroll (cheap DOM geometry, no canvas paint).
   */
  updateOverlay(): void {
    const overlayHost = this.host.overlayHost;
    overlayHost.clearFind();
    const ws = this.host.worksheet();
    if (!ws) return;
    const cs = this.host.scale();
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;
    const frozen = getGridGeometryForWorksheet(ws).roundedFrozenExtent(cs);
    const frozenBoundX = headerW + frozen.width;
    const frozenBoundY = headerH + frozen.height;

    // A match accent: same single-color → border + translucent fill derivation
    // the selection overlay uses. The active match uses a warm accent so it is
    // distinguishable from other hits and from the (blue) selection box.
    const colors = this.host.highlightColors();
    const other = findHighlightOverlayStyle(false, colors);
    const active = findHighlightOverlayStyle(true, colors);

    for (const hl of this.controller.sheetHighlights(this.host.currentSheet())) {
      const rect = this.host.cellRect(hl.row, hl.col);
      if (!rect) continue;
      let { x, y, w, h } = rect;
      // Clamp against headers + the frozen-pane boundary (scrollable cells that
      // scrolled behind the frozen area are clipped there), mirroring the
      // selection overlay so a highlight never spills over fixed regions.
      if (x < headerW) { w -= headerW - x; x = headerW; }
      if (y < headerH) { h -= headerH - y; y = headerH; }
      if (hl.col > freezeCols && x < frozenBoundX) { w -= frozenBoundX - x; x = frozenBoundX; }
      if (hl.row > freezeRows && y < frozenBoundY) { h -= frozenBoundY - y; y = frozenBoundY; }
      if (w <= 0 || h <= 0) continue;
      const screenLeft = this.host.screenX(x, w);
      const { border, background } = hl.active ? active : other;
      const box = this.host.ownerDocument.createElement('div');
      box.style.cssText =
        `position:absolute;` +
        `left:${screenLeft}px;top:${y}px;width:${w}px;height:${h}px;` +
        `box-sizing:border-box;border:${border};background:${background};pointer-events:none;`;
      overlayHost.appendFind(box);
    }
  }

  /** Search every sheet and highlight the matches (see XlsxViewer.findText). */
  async find(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<XlsxMatchLocation>[]> {
    if (!this.host.workbook()) return [];
    const matches = await this.controller.find(query, opts);
    this.updateOverlay();
    return matches;
  }

  /** Move to the next match (wrap-around). */
  async next(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this.activateMatch(this.controller.next());
  }

  /** Move to the previous match (wrap-around). */
  async prev(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this.activateMatch(this.controller.prev());
  }

  /** Clear all highlights and reset the find state. */
  clear(): void {
    this.controller.invalidate();
    this.updateOverlay();
  }

  private async activateMatch(
    match: FindMatch<XlsxMatchLocation> | null,
  ): Promise<FindMatch<XlsxMatchLocation> | null> {
    if (!match) {
      this.updateOverlay();
      return null;
    }
    const { sheet, row, col } = match.location;
    if (sheet !== this.host.currentSheet()) {
      // showSheet resets scroll/selection and re-renders; the find state (and so
      // the highlights) survive because they live on the controller, not the
      // sheet. updateOverlay runs after the sheet switch below.
      await this.host.goToSheet(sheet);
    }
    this.host.scrollCellIntoView(row, col);
    // Scrolling schedules a coalesced render; draw the highlights now so the
    // active box is visible immediately without waiting a frame.
    this.updateOverlay();
    return match;
  }

  /** Teardown: drop the find state so a stale findNext()/findPrev() after
   *  teardown returns null instead of a match pointing into a dead viewer. */
  destroy(): void {
    this.controller.invalidate();
  }
}
