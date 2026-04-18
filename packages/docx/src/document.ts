import type { Document, RenderPageOptions } from './types';
import { renderDocumentToCanvas } from './renderer';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';

export class DocxDocument {
  private _document: Document | null = null;

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

  private async _parse(buffer: ArrayBuffer): Promise<void> {
    const wasmModule = await import('./wasm/docx_parser.js');
    await wasmModule.default(new URL(wasmAssetUrl, location.href).href);
    const bytes = new Uint8Array(buffer);
    const json = wasmModule.parse_docx(bytes);
    const parsed = JSON.parse(json);
    if (parsed.error) throw new Error(`Parse error: ${parsed.error}`);
    this._document = parsed as Document;
  }

  get pageCount(): number {
    if (!this._document) return 0;
    let pages = 1;
    for (const el of this._document.body) {
      if (el.type === 'pageBreak') pages++;
    }
    return pages;
  }

  get document(): Document {
    if (!this._document) throw new Error('Document not loaded');
    return this._document;
  }

  renderPage(
    target: HTMLCanvasElement | OffscreenCanvas,
    pageIndex: number,
    opts: RenderPageOptions = {},
  ): Promise<void> {
    if (!this._document) throw new Error('Document not loaded');
    return renderDocumentToCanvas(this._document, target, pageIndex, {
      ...opts,
      totalPages: this.pageCount,
    });
  }
}
