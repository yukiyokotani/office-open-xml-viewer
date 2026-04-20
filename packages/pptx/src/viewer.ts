import type { RenderOptions } from './renderer';
import { PptxPresentation } from './presentation';

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
}

/**
 * Opinionated single-canvas PPTX viewer.
 *
 * Creates a <canvas> element, appends it to the provided container, and manages
 * slide navigation.
 *
 * For custom layouts (multi-canvas, thumbnails, scroll view) use PptxPresentation directly.
 */
export class PptxViewer {
  private readonly canvas: HTMLCanvasElement;
  private engine: PptxPresentation | null = null;
  private readonly opts: PptxViewerOptions;
  private currentSlide = 0;

  constructor(container: HTMLElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;

    this.canvas = document.createElement('canvas');
    this.canvas.style.display = 'block';
    this.canvas.style.maxWidth = '100%';
    this.canvas.style.height = 'auto';
    container.appendChild(this.canvas);
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

    try {
      await this.engine.renderSlide(this.canvas, this.currentSlide, { width: targetWidth, dpr });
      this.opts.onSlideChange?.(this.currentSlide, this.slideCount);
    } catch (err) {
      this.opts.onError?.(err instanceof Error ? err : new Error(String(err)));
    }
  }

  /** Clean up the viewer and terminate the background worker. */
  destroy(): void {
    this.engine?.destroy();
    this.canvas.remove();
  }
}
