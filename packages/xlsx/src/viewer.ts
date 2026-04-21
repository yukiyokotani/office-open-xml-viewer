import { XlsxWorkbook } from './workbook.js';
import type { ViewportRange, Worksheet } from './types.js';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx } from './renderer.js';

const TAB_BAR_H = 30;

export interface XlsxViewerOptions {
  /** Scale factor for cell/header dimensions (default 1). 0.5 = half size. */
  cellScale?: number;
  onReady?: (sheetNames: string[]) => void;
  onSheetChange?: (index: number, name: string) => void;
  onError?: (err: Error) => void;
}

export class XlsxViewer {
  private wb: XlsxWorkbook;
  private canvas: HTMLCanvasElement;
  private canvasArea: HTMLDivElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private tabBar: HTMLDivElement;
  private tabs: HTMLButtonElement[] = [];
  private currentSheet = 0;
  private currentWorksheet: Worksheet | null = null;
  private opts: XlsxViewerOptions;
  private resizeObserver: ResizeObserver | null = null;

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    this.opts = opts;
    this.wb = new XlsxWorkbook();

    const wrapper = document.createElement('div');
    wrapper.style.cssText =
      `position:relative;width:100%;height:100%;` +
      `border:1px solid #c8ccd0;background:#fff;box-sizing:border-box;font-family:sans-serif;display:flex;flex-direction:column;`;

    this.canvasArea = document.createElement('div');
    this.canvasArea.style.cssText = `position:relative;flex:1;min-height:0;overflow:hidden;`;

    this.canvas = document.createElement('canvas');
    this.canvas.style.cssText = `position:absolute;top:0;left:0;z-index:0;display:block;`;

    this.scrollHost = document.createElement('div');
    this.scrollHost.style.cssText = `position:absolute;inset:0;overflow:auto;z-index:1;background:transparent;`;

    this.spacer = document.createElement('div');
    this.spacer.style.cssText = `position:absolute;top:0;left:0;pointer-events:none;`;
    this.scrollHost.appendChild(this.spacer);

    this.canvasArea.appendChild(this.canvas);
    this.canvasArea.appendChild(this.scrollHost);

    this.tabBar = document.createElement('div');
    this.tabBar.style.cssText =
      `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
      `background:#f0f0f0;border-top:1px solid #c8ccd0;` +
      `overflow-x:auto;overflow-y:hidden;padding:0 4px;gap:1px;scrollbar-width:none;`;
    const style = document.createElement('style');
    style.textContent = `.xlsx-tab-bar::-webkit-scrollbar{display:none}`;
    document.head.appendChild(style);
    this.tabBar.classList.add('xlsx-tab-bar');

    wrapper.appendChild(this.canvasArea);
    wrapper.appendChild(this.tabBar);
    container.appendChild(wrapper);

    this.scrollHost.addEventListener('scroll', () => this.renderCurrentSheet());

    // Re-render whenever the canvas area changes size
    this.resizeObserver = new ResizeObserver(() => this.renderCurrentSheet());
    this.resizeObserver.observe(this.canvasArea);
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    try {
      await this.wb.load(source);
      this.buildTabs();
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
    this.updateTabActive(index);
    this.currentWorksheet = await this.wb.getWorksheet(index);
    this.updateSpacerSize(this.currentWorksheet);
    await this.renderCurrentSheet();
    this.opts.onSheetChange?.(index, this.wb.sheetNames[index] ?? '');
  }

  private buildTabs(): void {
    this.tabBar.innerHTML = '';
    this.tabs = [];
    this.wb.sheetNames.forEach((name, i) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabStyle(false);
      btn.addEventListener('click', () => this.showSheet(i));
      this.tabBar.appendChild(btn);
      this.tabs.push(btn);
    });
  }

  private updateTabActive(index: number): void {
    this.tabs.forEach((btn, i) => {
      btn.style.cssText = this.tabStyle(i === index);
    });
    this.tabs[index]?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  private tabStyle(active: boolean): string {
    // Active tab renders taller than inactive so the selected sheet draws the
    // eye. Tabs align to flex-end, so shorter inactive tabs sit lower and the
    // active tab sticks up. Font size also bumps a hair on active.
    const activeH = TAB_BAR_H - 2;
    const inactiveH = TAB_BAR_H - 5;
    const base =
      `display:inline-block;padding:0 14px;` +
      `border:1px solid #c8ccd0;border-bottom:none;border-radius:3px 3px 0 0;` +
      `cursor:pointer;white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;` +
      `outline:none;box-sizing:border-box;`;
    return active
      ? base +
        `height:${activeH}px;font-size:13px;` +
        `background:#fff;color:#000;border-bottom:1px solid #fff;font-weight:600;position:relative;top:1px;`
      : base +
        `height:${inactiveH}px;font-size:11px;` +
        `background:#e0e0e0;color:#555;`;
  }

  private updateSpacerSize(ws: Worksheet): void {
    const cs = this.opts.cellScale ?? 1;
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Compute frozen area pixel size
    let frozenW = 0;
    for (let c = 1; c <= freezeCols; c++) {
      frozenW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth);
    }
    let frozenH = 0;
    for (let r = 1; r <= freezeRows; r++) {
      frozenH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    }

    // Find actual scrollable data extent
    let maxRow = Math.max(50, freezeRows);
    let maxCol = Math.max(26, freezeCols);
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }
    maxRow += 30;
    maxCol += 10;

    // Spacer = header + frozen area + scrollable area (all in logical px, then scale)
    let totalW = HEADER_W + frozenW;
    for (let c = freezeCols + 1; c <= maxCol; c++) {
      totalW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth);
    }
    let totalH = HEADER_H + frozenH;
    for (let r = freezeRows + 1; r <= maxRow; r++) {
      totalH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    }

    this.spacer.style.width = `${Math.round(totalW * cs)}px`;
    this.spacer.style.height = `${Math.round(totalH * cs)}px`;
  }

  private async renderCurrentSheet(): Promise<void> {
    if (!this.currentWorksheet) return;
    const ws = this.currentWorksheet;
    const w = this.canvasArea.clientWidth;
    const h = this.canvasArea.clientHeight;
    if (w <= 0 || h <= 0) return;

    const cs = this.opts.cellScale ?? 1;
    const dpr = window.devicePixelRatio ?? 1;

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Compute frozen area in logical (unscaled) pixels
    let frozenW = 0;
    for (let c = 1; c <= freezeCols; c++) {
      frozenW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth);
    }
    let frozenH = 0;
    for (let r = 1; r <= freezeRows; r++) {
      frozenH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    }

    // DOM scrollLeft/scrollTop are in scaled (physical) CSS pixels.
    // Convert to logical pixels for cell-finding by dividing by cs.
    const logicalScrollX = this.scrollHost.scrollLeft / cs;
    const logicalScrollY = this.scrollHost.scrollTop / cs;

    // Find startCol in logical pixel space
    let startCol = freezeCols + 1;
    let xAcc = 0;
    let offsetX = 0;
    while (true) {
      const cw = colWidthToPx(ws.colWidths[startCol] ?? ws.defaultColWidth);
      if (xAcc + cw > logicalScrollX) { offsetX = logicalScrollX - xAcc; break; }
      xAcc += cw;
      startCol++;
      if (startCol > 16384) break;
    }

    // Find startRow in logical pixel space
    let startRow = freezeRows + 1;
    let yAcc = 0;
    let offsetY = 0;
    while (true) {
      const rh = rowHeightToPx(ws.rowHeights[startRow] ?? ws.defaultRowHeight);
      if (yAcc + rh > logicalScrollY) { offsetY = logicalScrollY - yAcc; break; }
      yAcc += rh;
      startRow++;
      if (startRow > 1048576) break;
    }

    // Effective scrollable area in logical pixels (canvas / cs - headers - frozen)
    const cellW = w / cs - HEADER_W - frozenW;
    const cellH = h / cs - HEADER_H - frozenH;

    // Compute exact number of visible columns by walking actual widths (+ 2 buffer)
    let cols = 0;
    { let xAcc = -offsetX; let c = startCol;
      while (xAcc < cellW + offsetX && c <= 16384) {
        xAcc += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth); cols++; c++;
      }
      cols += 2;
    }
    // Compute exact number of visible rows by walking actual heights (+ 2 buffer)
    let rows = 0;
    { let yAcc = -offsetY; let r = startRow;
      while (yAcc < cellH + offsetY && r <= 1048576) {
        yAcc += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight); rows++; r++;
      }
      rows += 2;
    }

    const viewport: ViewportRange = { row: startRow, col: startCol, rows, cols };

    await this.wb.renderViewport(this.canvas, this.currentSheet, viewport, {
      width: w,
      height: h,
      dpr,
      cellScale: cs,
      scrollOffsetX: offsetX,
      scrollOffsetY: offsetY,
      freezeRows,
      freezeCols,
    });
  }

  get sheetNames(): string[] {
    return this.wb.sheetNames;
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    this.wb.destroy();
  }
}
