import { DocxDocument } from './document';
import type { RenderPageOptions } from './types';

export interface DocxViewerOptions extends RenderPageOptions {
  container?: HTMLElement;
}

export class DocxViewer {
  private _doc: DocxDocument | null = null;
  private _currentPage = 0;
  private _canvas: HTMLCanvasElement;
  private _opts: DocxViewerOptions;

  constructor(canvas: HTMLCanvasElement, opts: DocxViewerOptions = {}) {
    this._canvas = canvas;
    this._opts = opts;
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    this._doc = await DocxDocument.load(source);
    this._currentPage = 0;
    this._render();
  }

  get pageCount(): number {
    return this._doc?.pageCount ?? 0;
  }

  get currentPage(): number {
    return this._currentPage;
  }

  goToPage(index: number): void {
    if (!this._doc) return;
    const clamped = Math.max(0, Math.min(index, this.pageCount - 1));
    this._currentPage = clamped;
    this._render();
  }

  nextPage(): void { this.goToPage(this._currentPage + 1); }
  prevPage(): void { this.goToPage(this._currentPage - 1); }

  private _render(): void {
    if (!this._doc) return;
    this._doc.renderPage(this._canvas, this._currentPage, this._opts);
  }
}
