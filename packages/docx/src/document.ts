import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import { preloadGoogleFonts, type FontPreloadEntry, type LoadOptions as CoreLoadOptions } from '@silurus/ooxml-core';
import type { BodyElement, Document, RenderPageOptions, WorkerResponse } from './types';
import { computePages, renderDocumentToCanvas } from './renderer';

/** Theme-referenced typefaces commonly used by DOCX templates. Mirrors the
 *  PPTX map — these are the well-known free webfont alternatives Microsoft
 *  Office templates pull from. Substitutes that diverge from the requested
 *  family name (Calibri → Carlito, Cambria → Caladea) include
 *  `loadFamily` so the FontFaceSet load is driven against the substitute. */
const DOCX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
};

/** Options for {@link DocxDocument.load}. Re-exports the shared
 *  `LoadOptions` shape from `@silurus/ooxml-core`. */
export type LoadOptions = CoreLoadOptions;

export class DocxDocument {
  private _document: Document | null = null;
  private _pages: BodyElement[][] | null = null;
  private _worker: Worker;

  private constructor() {
    this._worker = new InlineWorker();
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._worker.postMessage({ type: 'init', wasmUrl });
  }

  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<DocxDocument> {
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
    if (opts.useGoogleFonts) {
      await preloadGoogleFonts(
        [doc._document?.majorFont, doc._document?.minorFont],
        DOCX_GOOGLE_FONTS,
      );
    }
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
