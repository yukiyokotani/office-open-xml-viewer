import { XlsxWorkbook } from './workbook.js';
import type { ViewportRange } from './types.js';

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
  private currentSheet = 0;
  private opts: XlsxViewerOptions;

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    this.opts = opts;
    this.wb = new XlsxWorkbook();

    this.canvas = document.createElement('canvas');
    const w = opts.width ?? 1200;
    const h = opts.height ?? 600;
    this.canvas.style.cssText = `width:${w}px;height:${h}px;display:block;`;
    container.appendChild(this.canvas);
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
    const w = this.opts.width ?? 1200;
    const h = this.opts.height ?? 600;
    const dpr = typeof window !== 'undefined' ? window.devicePixelRatio : 1;

    // Compute how many rows/cols fit the canvas at default sizes
    const COL_W_PX = 8.43 * 7;  // default col width in px
    const ROW_H_PX = 15 * 1.333; // default row height in px
    const cols = Math.ceil(w / COL_W_PX) + 2;
    const rows = Math.ceil(h / ROW_H_PX) + 2;

    const viewport: ViewportRange = { row: 1, col: 1, rows, cols };

    await this.wb.renderViewport(this.canvas, index, viewport, { width: w, height: h, dpr });
    this.opts.onSheetChange?.(index, this.wb.sheetNames[index] ?? '');
  }

  get sheetNames(): string[] {
    return this.wb.sheetNames;
  }

  destroy(): void {
    this.wb.destroy();
  }
}
