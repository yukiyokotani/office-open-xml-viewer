import type { RenderOptions, TextRunInfo } from './renderer';
import { PptxPresentation } from './presentation';
import type { PresentationHandle } from './presentation-handle';

export interface PptxViewerOptions extends RenderOptions {
  /** Called when a slide finishes rendering */
  onSlideChange?: (index: number, total: number) => void;
  /** Called on parse or render errors */
  onError?: (err: Error) => void;
  /**
   * Opt in to loading theme-declared webfonts from Google Fonts. Off by
   * default — see {@link PptxPresentation.load} for privacy implications.
   */
  useGoogleFonts?: boolean;
  /**
   * Enable interactive audio/video playback. When true, slides are rendered
   * via {@link PptxPresentation.presentSlide} so media elements become
   * clickable and the viewer draws its own play/pause chrome. When false
   * (default) the viewer renders a static slide with a non-interactive play
   * badge over media posters.
   */
  enableMediaPlayback?: boolean;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on slide content.
   */
  enableTextSelection?: boolean;
}

/**
 * Opinionated single-canvas PPTX viewer.
 *
 * Accepts a caller-supplied `<canvas>` element and wraps it in a positioned
 * container for the optional text-selection overlay.  The wrapper is inserted
 * into the canvas's existing parent (reparent), so the canvas stays at its
 * original position in the DOM.
 *
 * For custom layouts (multi-canvas, thumbnails, scroll view) use PptxPresentation directly.
 */
export class PptxViewer {
  private readonly canvas: HTMLCanvasElement;
  private readonly wrapper: HTMLDivElement;
  private textLayer: HTMLDivElement | null = null;
  private engine: PptxPresentation | null = null;
  private readonly opts: PptxViewerOptions;
  private currentSlide = 0;
  private handle: PresentationHandle | null = null;

  constructor(canvas: HTMLCanvasElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;
    this.canvas = canvas;

    const parent = canvas.parentElement;
    this.wrapper = document.createElement('div');
    // vertical-align:top removes the inline-block baseline descender gap that
    // otherwise lets the host container's background show through below the
    // canvas (~6 px on default font metrics).
    this.wrapper.style.cssText = 'position:relative;display:inline-block;vertical-align:top;';
    if (parent) parent.insertBefore(this.wrapper, canvas);
    this.wrapper.appendChild(canvas);

    if (opts.enableTextSelection) {
      this.textLayer = document.createElement('div');
      this.textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      this.wrapper.appendChild(this.textLayer);
    }
  }

  /** Load a PPTX from URL or ArrayBuffer and render the first slide. */
  async load(source: string | ArrayBuffer): Promise<void> {
    try {
      this.engine = await PptxPresentation.load(source, {
        useGoogleFonts: this.opts.useGoogleFonts,
      });
      this.currentSlide = 0;
      await this.renderCurrentSlide();
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      this.opts.onError?.(e);
      throw e;
    }
  }

  /** Navigate to a specific slide (0-indexed). */
  async goToSlide(index: number): Promise<void> {
    if (!this.engine || this.slideCount === 0) return;
    this.currentSlide = Math.max(0, Math.min(index, this.slideCount - 1));
    await this.renderCurrentSlide();
  }

  async nextSlide(): Promise<void> {
    await this.goToSlide(this.currentSlide + 1);
  }

  async prevSlide(): Promise<void> {
    await this.goToSlide(this.currentSlide - 1);
  }

  get slideIndex(): number { return this.currentSlide; }
  get slideCount(): number { return this.engine?.slideCount ?? 0; }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement { return this.canvas; }

  private async renderCurrentSlide(): Promise<void> {
    if (!this.engine) return;
    const targetWidth = this.opts.width ?? (this.canvas.offsetWidth || 960);
    const dpr = this.opts.dpr ?? (window.devicePixelRatio || 1);

    const scale = targetWidth / this.engine.slideWidth;
    const cssHeight = Math.round(this.engine.slideHeight * scale);
    this.canvas.style.width = `${targetWidth}px`;
    this.canvas.style.height = `${cssHeight}px`;

    this.handle?.dispose();
    this.handle = null;

    const runs: TextRunInfo[] = [];
    const onTextRun = this.textLayer ? (r: TextRunInfo) => runs.push(r) : undefined;

    try {
      if (this.opts.enableMediaPlayback) {
        this.handle = await this.engine.presentSlide(this.canvas, this.currentSlide, {
          width: targetWidth,
          dpr,
        });
      } else {
        await this.engine.renderSlide(this.canvas, this.currentSlide, { width: targetWidth, dpr, onTextRun });
      }
      this.opts.onSlideChange?.(this.currentSlide, this.slideCount);
    } catch (err) {
      this.opts.onError?.(err instanceof Error ? err : new Error(String(err)));
    }

    if (this.textLayer) {
      this._buildTextLayer(runs, targetWidth, cssHeight);
    }
  }

  private _buildTextLayer(runs: TextRunInfo[], cssWidth: number, cssHeight: number): void {
    const layer = this.textLayer!;
    layer.innerHTML = '';
    layer.style.width = `${cssWidth}px`;
    layer.style.height = `${cssHeight}px`;

    // Group runs by shape (same shapeX/shapeY/rotation)
    type ShapeKey = string;
    const shapeMap = new Map<ShapeKey, { div: HTMLDivElement; x: number; y: number; w: number; h: number; rot: number }>();

    for (const run of runs) {
      const totalRot = run.rotation + (run.textBodyRotation ?? 0);
      const key = `${run.shapeX},${run.shapeY},${run.shapeW},${run.shapeH},${totalRot}`;
      if (!shapeMap.has(key)) {
        const div = document.createElement('div');
        div.style.cssText =
          `position:absolute;` +
          `left:${run.shapeX}px;top:${run.shapeY}px;` +
          `width:${run.shapeW}px;height:${run.shapeH}px;` +
          `pointer-events:all;overflow:hidden;`;
        if (totalRot !== 0) {
          div.style.transformOrigin = 'center center';
          div.style.transform = `rotate(${totalRot}deg)`;
        }
        shapeMap.set(key, { div, x: run.shapeX, y: run.shapeY, w: run.shapeW, h: run.shapeH, rot: totalRot });
        layer.appendChild(div);
      }

      const shape = shapeMap.get(key)!;
      const span = document.createElement('span');
      span.textContent = run.text;
      // The `font` shorthand must precede `line-height` because the shorthand
      // resets `line-height` to `normal`. Reset `letter-spacing` so a parent
      // CSS rule cannot drift the trailing edge of the selection. Kerning /
      // ligatures are left at the browser default ('auto') because canvas
      // `measureText` / `fillText` also apply them by default — forcing them
      // off here would make the span wider than the drawn text.
      span.style.cssText =
        `position:absolute;` +
        `left:${run.inShapeX}px;top:${run.inShapeY}px;` +
        `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
        `white-space:pre;color:transparent;cursor:text;`;
      shape.div.appendChild(span);
    }
  }

  /** Clean up the viewer and terminate the background worker. */
  destroy(): void {
    this.handle?.dispose();
    this.handle = null;
    this.engine?.destroy();
    this.wrapper.remove();
  }
}
