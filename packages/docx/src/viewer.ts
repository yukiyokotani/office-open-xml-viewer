import { DocxDocument } from './document';
import type { RenderPageOptions } from './types';
import type { DocxTextRunInfo } from './renderer';

export interface DocxViewerOptions extends RenderPageOptions {
  container?: HTMLElement;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on document content.
   */
  enableTextSelection?: boolean;
}

export class DocxViewer {
  private _doc: DocxDocument | null = null;
  private _currentPage = 0;
  private _canvas: HTMLCanvasElement;
  private _wrapper: HTMLDivElement;
  private _textLayer: HTMLDivElement | null = null;
  private _opts: DocxViewerOptions;

  constructor(canvas: HTMLCanvasElement, opts: DocxViewerOptions = {}) {
    this._canvas = canvas;
    this._opts = opts;

    // Wrap canvas in a positioned container for the optional text layer overlay
    const parent = canvas.parentElement;
    this._wrapper = document.createElement('div');
    this._wrapper.style.cssText = 'position:relative;display:inline-block;';
    if (parent) {
      parent.insertBefore(this._wrapper, canvas);
    }
    this._wrapper.appendChild(canvas);

    if (opts.enableTextSelection) {
      this._textLayer = document.createElement('div');
      this._textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      this._wrapper.appendChild(this._textLayer);
    }
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    this._doc = await DocxDocument.load(source);
    this._currentPage = 0;
    await this._render();
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

  private async _render(): Promise<void> {
    if (!this._doc) return;
    const runs: DocxTextRunInfo[] = [];
    const onTextRun = this._textLayer ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
    await this._doc.renderPage(this._canvas, this._currentPage, { ...this._opts, onTextRun });
    if (this._textLayer) {
      this._buildTextLayer(runs);
    }
  }

  private _buildTextLayer(runs: DocxTextRunInfo[]): void {
    const layer = this._textLayer!;
    layer.innerHTML = '';
    layer.style.width = `${this._canvas.style.width || this._canvas.width + 'px'}`;
    layer.style.height = `${this._canvas.style.height || this._canvas.height + 'px'}`;

    for (const run of runs) {
      const span = document.createElement('span');
      span.textContent = run.text;
      span.style.cssText =
        `position:absolute;` +
        `left:${run.x}px;top:${run.y}px;` +
        `font-size:${run.fontSize}px;line-height:${run.h}px;` +
        `white-space:pre;color:transparent;cursor:text;pointer-events:all;`;
      layer.appendChild(span);
    }
  }
}
