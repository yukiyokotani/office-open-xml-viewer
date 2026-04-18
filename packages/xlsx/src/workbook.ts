import InlineWorker from './worker.ts?worker&inline';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions, WorkerResponse } from './types.js';
import { renderViewport } from './renderer.js';

export class XlsxWorkbook {
  private worker: Worker;
  private parsedWorkbook: ParsedWorkbook | null = null;
  private sheetCache = new Map<number, Worksheet>();
  private rawData: ArrayBuffer | null = null;

  constructor() {
    this.worker = new InlineWorker();
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    const data =
      typeof source === 'string'
        ? await fetch(source).then((r) => r.arrayBuffer())
        : source;
    this.rawData = data;
    this.parsedWorkbook = await this.sendMessage({ type: 'parse', data: data.slice(0) }) as ParsedWorkbook;
  }

  get sheetNames(): string[] {
    return this.parsedWorkbook?.workbook.sheets.map((s) => s.name) ?? [];
  }

  get sheetCount(): number {
    return this.parsedWorkbook?.workbook.sheets.length ?? 0;
  }

  async getWorksheet(sheetIndex: number): Promise<Worksheet> {
    if (this.sheetCache.has(sheetIndex)) {
      return this.sheetCache.get(sheetIndex)!;
    }
    if (!this.parsedWorkbook || !this.rawData) {
      throw new Error('Workbook not loaded');
    }
    const sheetMeta = this.parsedWorkbook.workbook.sheets[sheetIndex];
    if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);

    const ws = await this.sendMessage({
      type: 'parseSheet',
      data: this.rawData.slice(0),
      sheetIndex,
      sheetName: sheetMeta.name,
    }) as Worksheet;
    this.sheetCache.set(sheetIndex, ws);
    return ws;
  }

  async renderViewport(
    target: HTMLCanvasElement | OffscreenCanvas,
    sheetIndex: number,
    viewport: ViewportRange,
    opts: RenderViewportOptions = {},
  ): Promise<void> {
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    const ws = await this.getWorksheet(sheetIndex);
    const styles = this.parsedWorkbook.styles;

    const dpr = opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio : 1);
    const rawW = target instanceof HTMLCanvasElement ? (target.clientWidth || 800) : target.width;
    const rawH = target instanceof HTMLCanvasElement ? (target.clientHeight || 600) : target.height;
    const width = opts.width ?? rawW;
    const height = opts.height ?? rawH;

    target.width = Math.round(width * dpr);
    target.height = Math.round(height * dpr);
    // Set CSS display size so the browser renders at 1:1 device pixels (no browser-level scaling).
    // Without this, canvas.width=2400 on a DPR=2 display causes the canvas to be laid out at
    // 2400 CSS px, making all content appear blurry when viewed in a 1200 CSS px container.
    if (target instanceof HTMLCanvasElement) {
      target.style.width = `${width}px`;
      target.style.height = `${height}px`;
    }

    const ctx = (target as HTMLCanvasElement).getContext('2d') as CanvasRenderingContext2D;
    ctx.scale(dpr, dpr);

    renderViewport(ctx, ws, styles, viewport, { ...opts, dpr });
  }

  destroy(): void {
    this.worker.terminate();
  }

  private sendMessage(req: object): Promise<ParsedWorkbook | Worksheet> {
    return new Promise((resolve, reject) => {
      const handler = (e: MessageEvent<WorkerResponse>) => {
        this.worker.removeEventListener('message', handler);
        if (e.data.type === 'error') {
          reject(new Error(e.data.message));
        } else if (e.data.type === 'parsed') {
          resolve(e.data.workbook);
        } else if (e.data.type === 'parsedSheet') {
          resolve(e.data.worksheet);
        }
      };
      this.worker.addEventListener('message', handler);
      this.worker.postMessage(req);
    });
  }
}
