import { XlsxWorkbook } from './workbook.js';
import type { ViewportRange, Worksheet } from './types.js';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx } from './renderer.js';

export interface XlsxViewerOptions {
  width?: number;
  height?: number;
  onReady?: (sheetNames: string[]) => void;
  onSheetChange?: (index: number, name: string) => void;
  onError?: (err: Error) => void;
}

export class XlsxViewer {
  private wb: XlsxWorkbook;
  private canvas: HTMLCanvasElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private currentSheet = 0;
  private currentWorksheet: Worksheet | null = null;
  private opts: XlsxViewerOptions;

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    this.opts = opts;
    this.wb = new XlsxWorkbook();

    const w = opts.width ?? 1200;
    const h = opts.height ?? 600;

    // Outer wrapper: clips overflow and positions children
    const wrapper = document.createElement('div');
    wrapper.style.cssText = `position:relative;width:${w}px;height:${h}px;overflow:hidden;border:1px solid #c8ccd0;background:#fff;`;

    // Canvas rendered underneath (z-index:0)
    this.canvas = document.createElement('canvas');
    this.canvas.style.cssText = `position:absolute;top:0;left:0;z-index:0;display:block;`;

    // Scroll host on top (z-index:1). Transparent background so canvas shows through.
    // Intercepts scroll events; the spacer inside defines the virtual document size.
    this.scrollHost = document.createElement('div');
    this.scrollHost.style.cssText = `position:absolute;inset:0;overflow:auto;z-index:1;background:transparent;`;

    // Spacer defines the virtual scroll range.
    // Width/height are updated after the worksheet loads.
    this.spacer = document.createElement('div');
    this.spacer.style.cssText = `position:absolute;top:0;left:0;width:${w}px;height:${h}px;pointer-events:none;`;
    this.scrollHost.appendChild(this.spacer);

    wrapper.appendChild(this.canvas);
    wrapper.appendChild(this.scrollHost);
    container.appendChild(wrapper);

    this.scrollHost.addEventListener('scroll', () => this.renderCurrentSheet());
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    try {
      await this.wb.load(source);
      this.opts.onReady?.(this.wb.sheetNames);
      await this.showSheet(0);
    } catch (err) {
      this.opts.onError?.(err instanceof Error ? err : new Error(String(err)));
    }
  }

  async showSheet(index: number): Promise<void> {
    this.currentSheet = index;
    this.scrollHost.scrollLeft = 0;
    this.scrollHost.scrollTop = 0;
    this.currentWorksheet = await this.wb.getWorksheet(index);
    this.updateSpacerSize(this.currentWorksheet);
    await this.renderCurrentSheet();
    this.opts.onSheetChange?.(index, this.wb.sheetNames[index] ?? '');
  }

  private updateSpacerSize(ws: Worksheet): void {
    // Find actual data extent from rows
    let maxRow = 50;
    let maxCol = 26;
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }
    maxRow = maxRow + 30;
    maxCol = maxCol + 10;

    // Compute total cell data dimensions (not including headers, but add header offset for spacer)
    let totalW = HEADER_W;
    for (let c = 1; c <= maxCol; c++) {
      totalW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth);
    }
    let totalH = HEADER_H;
    for (let r = 1; r <= maxRow; r++) {
      totalH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    }

    this.spacer.style.width = `${totalW}px`;
    this.spacer.style.height = `${totalH}px`;
  }

  private async renderCurrentSheet(): Promise<void> {
    if (!this.currentWorksheet) return;
    const ws = this.currentWorksheet;
    const w = this.opts.width ?? 1200;
    const h = this.opts.height ?? 600;
    const dpr = window.devicePixelRatio ?? 1;

    // Scroll position = pixel offset into cell data area (from start of col 1 / row 1)
    const scrollX = this.scrollHost.scrollLeft;
    const scrollY = this.scrollHost.scrollTop;

    // Find startCol and pixel offset within it
    let startCol = 1;
    let xAcc = 0;
    let offsetX = 0;
    while (true) {
      const cw = colWidthToPx(ws.colWidths[startCol] ?? ws.defaultColWidth);
      if (xAcc + cw > scrollX) {
        offsetX = scrollX - xAcc;
        break;
      }
      xAcc += cw;
      startCol++;
      if (startCol > 16384) break;
    }

    // Find startRow and pixel offset within it
    let startRow = 1;
    let yAcc = 0;
    let offsetY = 0;
    while (true) {
      const rh = rowHeightToPx(ws.rowHeights[startRow] ?? ws.defaultRowHeight);
      if (yAcc + rh > scrollY) {
        offsetY = scrollY - yAcc;
        break;
      }
      yAcc += rh;
      startRow++;
      if (startRow > 1048576) break;
    }

    // Effective cell area size (canvas minus headers)
    const cellW = w - HEADER_W;
    const cellH = h - HEADER_H;
    // Estimate number of rows/cols needed (+2 for partial cells at edges)
    const avgCW = colWidthToPx(ws.defaultColWidth);
    const avgRH = rowHeightToPx(ws.defaultRowHeight);
    const cols = Math.ceil((cellW + offsetX) / Math.max(avgCW, 1)) + 2;
    const rows = Math.ceil((cellH + offsetY) / Math.max(avgRH, 1)) + 2;

    const viewport: ViewportRange = { row: startRow, col: startCol, rows, cols };

    await this.wb.renderViewport(this.canvas, this.currentSheet, viewport, {
      width: w,
      height: h,
      dpr,
      scrollOffsetX: offsetX,
      scrollOffsetY: offsetY,
    });
  }

  get sheetNames(): string[] {
    return this.wb.sheetNames;
  }

  destroy(): void {
    this.wb.destroy();
  }
}
