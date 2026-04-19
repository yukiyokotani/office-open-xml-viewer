import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import type { BodyElement, Document, RenderPageOptions, WorkerResponse } from './types';
import { computePages, renderDocumentToCanvas } from './renderer';

export class DocxDocument {
  private _document: Document | null = null;
  private _pages: BodyElement[][] | null = null;
  private _worker: Worker;

  private constructor() {
    this._worker = new InlineWorker();
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._worker.postMessage({ type: 'init', wasmUrl });
  }

  static async load(source: string | ArrayBuffer): Promise<DocxDocument> {
    const doc = new DocxDocument();
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    await doc._parse(buffer);
    return doc;
  }

  private _parse(buffer: ArrayBuffer): Promise<void> {
    return new Promise((resolve, reject) => {
      const handler = (e: MessageEvent<WorkerResponse>) => {
        this._worker.removeEventListener('message', handler);
        if (e.data.type === 'error') {
          reject(new Error(e.data.message));
        } else if (e.data.type === 'parsed') {
          this._document = e.data.document;
          resolve();
        }
      };
      this._worker.addEventListener('message', handler);
      this._worker.postMessage({ type: 'parse', data: buffer }, [buffer]);
    });
  }

  destroy(): void {
    this._worker.terminate();
  }

  get pageCount(): number {
    if (!this._document) return 0;
    return this._getPages().length;
  }

  get document(): Document {
    if (!this._document) throw new Error('Document not loaded');
    return this._document;
  }

  private _getPages(): BodyElement[][] {
    if (this._pages) return this._pages;
    if (!this._document) return [];
    const measure = new OffscreenCanvas(1, 1);
    const ctx = measure.getContext('2d');
    if (!ctx) {
      this._pages = [this._document.body];
      return this._pages;
    }
    this._pages = computePages(this._document.body, this._document.section, ctx);
    return this._pages;
  }

  renderPage(
    target: HTMLCanvasElement | OffscreenCanvas,
    pageIndex: number,
    opts: RenderPageOptions = {},
  ): Promise<void> {
    if (!this._document) throw new Error('Document not loaded');
    const pages = this._getPages();
    return renderDocumentToCanvas(this._document, target, pageIndex, {
      ...opts,
      totalPages: pages.length,
      prebuiltPages: pages,
    });
  }
}
