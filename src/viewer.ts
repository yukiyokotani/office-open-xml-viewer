import type { Presentation, WorkerRequest, WorkerResponse } from './types';
import type { RenderOptions } from './renderer';
// Inline the worker so the library is self-contained (no separate worker file needed)
import InlineWorker from './worker.ts?worker&inline';
// Resolved by Vite at build time; passed to the worker so it can init WASM
// without relying on import.meta.url (which is blob: in inline workers).
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';

export interface PptxViewerOptions extends RenderOptions {
  /** Called when the viewer is ready to display slides */
  onReady?: () => void;
  /** Called when a slide finishes rendering */
  onSlideChange?: (index: number, total: number) => void;
  /** Called on parse or render errors */
  onError?: (err: Error) => void;
}

export class PptxViewer {
  private readonly canvas: HTMLCanvasElement;
  private readonly opts: PptxViewerOptions;
  private worker: Worker | null = null;
  private presentation: Presentation | null = null;
  private currentSlide = 0;
  private pendingCallbacks = new Map<
    number,
    { resolve: (p: Presentation) => void; reject: (e: Error) => void }
  >();
  private pendingRenderCallbacks = new Map<number, { resolve: () => void; reject: (e: Error) => void }>();
  private nextId = 1;
  private workerReady = false;
  private workerReadyCallbacks: Array<() => void> = [];
  private canvasTransferred = false;

  constructor(container: HTMLElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;

    this.canvas = document.createElement('canvas');
    this.canvas.style.display = 'block';
    this.canvas.style.maxWidth = '100%';
    this.canvas.style.height = 'auto';
    container.appendChild(this.canvas);

    this.initWorker();
  }

  private initWorker() {
    this.worker = new InlineWorker();
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this.worker.postMessage({ kind: 'init', wasmUrl } satisfies WorkerRequest);

    this.worker.onmessage = (e: MessageEvent<WorkerResponse>) => {
      const msg = e.data;

      if (msg.kind === 'ready') {
        // Transfer canvas control to the worker so it can render via OffscreenCanvas
        const offscreen = this.canvas.transferControlToOffscreen();
        const dpr = window.devicePixelRatio || 1;
        this.worker!.postMessage(
          { kind: 'transferCanvas', canvas: offscreen, dpr } satisfies WorkerRequest,
          [offscreen]
        );
        this.canvasTransferred = true;

        this.workerReady = true;
        for (const cb of this.workerReadyCallbacks) cb();
        this.workerReadyCallbacks = [];
        this.opts.onReady?.();
        return;
      }

      if (msg.kind === 'parsed') {
        const cb = this.pendingCallbacks.get(msg.id);
        if (cb) {
          this.pendingCallbacks.delete(msg.id);
          cb.resolve(msg.presentation);
        }
        return;
      }

      if (msg.kind === 'rendered') {
        const cb = this.pendingRenderCallbacks.get(msg.id);
        if (cb) {
          this.pendingRenderCallbacks.delete(msg.id);
          cb.resolve();
        }
        return;
      }

      if (msg.kind === 'error') {
        const cb = this.pendingCallbacks.get(msg.id) || null;
        const renderCb = this.pendingRenderCallbacks.get(msg.id) || null;
        const err = new Error(msg.message);
        if (cb) {
          this.pendingCallbacks.delete(msg.id);
          cb.reject(err);
        }
        if (renderCb) {
          this.pendingRenderCallbacks.delete(msg.id);
          renderCb.reject(err);
        }
        this.opts.onError?.(err);
      }
    };

    this.worker.onerror = (e) => {
      const err = new Error(e.message);
      this.opts.onError?.(err);
    };
  }

  private waitForWorker(): Promise<void> {
    if (this.workerReady) return Promise.resolve();
    return new Promise((resolve) => this.workerReadyCallbacks.push(resolve));
  }

  /** Load a PPTX file from an ArrayBuffer and render the first slide */
  async load(buffer: ArrayBuffer): Promise<void> {
    await this.waitForWorker();

    const id = this.nextId++;
    const presentation = await new Promise<Presentation>((resolve, reject) => {
      this.pendingCallbacks.set(id, { resolve, reject });
      const req: WorkerRequest = { kind: 'parse', id, buffer };
      this.worker!.postMessage(req, [buffer]);
    });

    console.log('[PptxViewer] parsed', {
      slideWidth: presentation.slideWidth,
      slideHeight: presentation.slideHeight,
      slideCount: presentation.slides.length,
      slide0elements: presentation.slides[0]?.elements.length ?? 0,
    });
    this.presentation = presentation;
    this.currentSlide = 0;
    await this.renderCurrentSlide();
  }

  /** Navigate to a specific slide (0-indexed) */
  async goToSlide(index: number): Promise<void> {
    if (!this.presentation) return;
    const clamped = Math.max(0, Math.min(index, this.presentation.slides.length - 1));
    this.currentSlide = clamped;
    await this.renderCurrentSlide();
  }

  async nextSlide(): Promise<void> {
    await this.goToSlide(this.currentSlide + 1);
  }

  async prevSlide(): Promise<void> {
    await this.goToSlide(this.currentSlide - 1);
  }

  get slideIndex(): number {
    return this.currentSlide;
  }

  get slideCount(): number {
    return this.presentation?.slides.length ?? 0;
  }

  private async renderCurrentSlide() {
    if (!this.presentation) return;

    // CSS sizing: set display width on the original canvas element (CSS still applies after transfer)
    const targetWidth = this.opts.width ?? (this.canvas.offsetWidth || 960);
    const scale = targetWidth / this.presentation.slideWidth;
    const canvasH = Math.round(this.presentation.slideHeight * scale);
    this.canvas.style.width = `${targetWidth}px`;
    // height:auto is set in constructor; this ensures the correct intrinsic aspect ratio is represented
    // by setting the CSS height explicitly from the slide dimensions.
    this.canvas.style.height = `${canvasH}px`;

    const id = this.nextId++;
    await new Promise<void>((resolve, reject) => {
      this.pendingRenderCallbacks.set(id, { resolve, reject });
      const req: WorkerRequest = {
        kind: 'render',
        id,
        slideIndex: this.currentSlide,
        targetWidth,
        defaultTextColor: this.presentation!.defaultTextColor,
      };
      this.worker!.postMessage(req);
    });

    this.opts.onSlideChange?.(this.currentSlide, this.presentation.slides.length);
  }

  /** Get the underlying <canvas> element */
  get canvasElement(): HTMLCanvasElement {
    return this.canvas;
  }

  /** Clean up the viewer and terminate the background worker */
  destroy() {
    this.worker?.terminate();
    this.worker = null;
    this.canvas.remove();
  }
}
