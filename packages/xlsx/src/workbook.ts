import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/xlsx_parser_bg.wasm?url';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions, WorkerResponse } from './types.js';
import { renderViewport } from './renderer.js';

/** Office font name → Google Fonts URL for a metric-compatible substitute.
 *  These pairings are the well-known Office substitutes that Microsoft and
 *  Google both publish and ship on Linux distributions:
 *    Calibri → Carlito (same metrics — advance widths, ascender / descender)
 *    Cambria → Caladea (likewise)
 *  Loading them on a system that lacks Calibri / Cambria (macOS, Linux
 *  without Office) keeps text width measurements close to Excel's. */
const OFFICE_FONT_FALLBACK_MAP: Record<string, string> = {
  'calibri': 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'cambria': 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap',
};

async function preloadOfficeFonts(fontNames: Iterable<string>): Promise<void> {
  if (typeof document === 'undefined') return;
  const seen = new Set<string>();
  const families: string[] = [];
  for (const name of fontNames) {
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const url = OFFICE_FONT_FALLBACK_MAP[key];
    if (!url) continue;
    if (document.querySelector(`link[href="${url}"]`)) continue;
    try {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = url;
      document.head.appendChild(link);
    } catch {
      // Network or DOM error — silently skip; renderer falls back to system fonts.
    }
    // We need the *substitute* family loaded into FontFaceSet, not the
    // requested Office name. Map calibri → Carlito, cambria → Caladea.
    families.push(key === 'calibri' ? 'Carlito' : key === 'cambria' ? 'Caladea' : name);
  }
  // Trigger explicit loads so Canvas measureText sees real glyph metrics
  // before the first render. @font-face declarations alone don't fetch.
  const loads: Promise<unknown>[] = [];
  for (const family of families) {
    for (const weight of ['400', '700']) {
      for (const style of ['normal', 'italic']) {
        loads.push(
          document.fonts.load(`${style} ${weight} 16px "${family}"`).catch(() => undefined),
        );
      }
    }
  }
  await Promise.race([
    Promise.all(loads).then(() => document.fonts.ready),
    new Promise<void>((resolve) => setTimeout(resolve, 3000)),
  ]);
}

/** Options for {@link XlsxWorkbook.load}. */
export interface LoadOptions {
  /**
   * Opt in to loading Office-font metric substitutes (Carlito for Calibri,
   * Caladea for Cambria) from Google Fonts (`fonts.googleapis.com`). Without
   * these the canvas falls back to system Arial / Helvetica which is
   * noticeably wider per character, so column layouts diverge from Excel.
   *
   * When enabled, end-user IP / User-Agent is sent to Google, which may have
   * privacy / GDPR implications for your application. Default `false` — host
   * the substitute webfonts yourself and reference them via `@font-face` in
   * your application CSS to avoid the third-party request.
   */
  useGoogleFonts?: boolean;
}

export class XlsxWorkbook {
  private worker: Worker;
  private parsedWorkbook: ParsedWorkbook | null = null;
  private sheetCache = new Map<number, Worksheet>();
  /** Cache of loaded images keyed by their data URL. Shared across sheets. */
  private imageCache = new Map<string, HTMLImageElement>();
  private rawData: ArrayBuffer | null = null;

  constructor() {
    this.worker = new InlineWorker();
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this.worker.postMessage({ type: 'init', wasmUrl });
  }

  async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<void> {
    const data =
      typeof source === 'string'
        ? await fetch(source).then((r) => r.arrayBuffer())
        : source;
    this.rawData = data;
    this.parsedWorkbook = await this.sendMessage({ type: 'parse', data: data.slice(0) }) as ParsedWorkbook;
    if (opts.useGoogleFonts) {
      // Walk every styled font in the workbook and queue Google Fonts
      // substitutes for any Office faces (Calibri → Carlito, Cambria →
      // Caladea). Documents that use only system fonts produce zero
      // network requests.
      const names = new Set<string>();
      for (const f of this.parsedWorkbook.styles?.fonts ?? []) {
        if (f.name) names.add(f.name);
      }
      await preloadOfficeFonts(names);
    }
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
    // Hot path: during scroll the worksheet is already cached. Skip the await
    // to keep the whole render in a single synchronous task so the browser
    // doesn't paint between the canvas clear (below) and the draw.
    const ws = this.sheetCache.get(sheetIndex) ?? await this.getWorksheet(sheetIndex);
    const styles = this.parsedWorkbook.styles;

    // ── Step 1: Preload any uncached image bitmaps BEFORE touching the canvas.
    //
    // Images can appear either as top-level twoCellAnchor `<xdr:pic>` (captured
    // in `ws.images`) or as a leaf inside an `<xdr:grpSp>` (captured as a
    // ShapeGeom with `type: 'image'`). We collect both so the renderer never
    // hits a missing bitmap during the synchronous draw pass.
    //
    // Doing this *before* the canvas resize is critical for scroll smoothness:
    // setting `canvas.width` wipes the canvas, and an `await` after that wipe
    // yields to the browser's paint cycle, causing a visible white flash on
    // every scroll frame. By awaiting first (and only when there's something
    // uncached), the whole resize+draw runs synchronously in a single tick and
    // the old frame stays visible until the new one is ready.
    const uncached: string[] = [];
    if (ws.images) {
      for (const img of ws.images) {
        if (!this.imageCache.has(img.dataUrl)) uncached.push(img.dataUrl);
      }
    }
    if (ws.shapeGroups) {
      for (const grp of ws.shapeGroups) {
        for (const shape of grp.shapes) {
          if (shape.geom.type === 'image' && !this.imageCache.has(shape.geom.dataUrl)) {
            uncached.push(shape.geom.dataUrl);
          }
        }
      }
    }
    if (uncached.length > 0) {
      await Promise.all(
        uncached.map(async (url) => {
          const el = new Image();
          el.src = url;
          await new Promise<void>((resolve, reject) => {
            el.onload = () => resolve();
            el.onerror = () => reject(new Error('image decode failed'));
          });
          this.imageCache.set(url, el);
        }),
      ).catch(() => { /* swallow image failures so the grid still renders */ });
    }

    // ── Step 2: Resize + draw, all synchronous from here.
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

    renderViewport(ctx, ws, styles, viewport, { ...opts, dpr, loadedImages: this.imageCache });
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
