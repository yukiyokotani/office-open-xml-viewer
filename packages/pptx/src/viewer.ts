import type { RenderOptions } from './renderer';
import { PptxPresentation } from './presentation';

export interface PptxViewerOptions extends RenderOptions {
  /** Called when the viewer is ready to display slides */
  onReady?: () => void;
  /** Called when a slide finishes rendering */
  onSlideChange?: (index: number, total: number) => void;
  /** Called on parse or render errors */
  onError?: (err: Error) => void;
}

/**
 * Opinionated single-canvas PPTX viewer.
 *
 * Creates a <canvas> element, appends it to the provided container, and manages
 * slide navigation. Public API is unchanged from the previous version.
 *
 * For custom layouts (multi-canvas, thumbnails, scroll view) use PptxPresentation directly.
 */
export class PptxViewer {
  private readonly canvas: HTMLCanvasElement;
  private readonly engine: PptxPresentation;
  private readonly opts: PptxViewerOptions;
  private currentSlide = 0;
  private loadedSlideCount = 0;

  constructor(container: HTMLElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;

    this.canvas = document.createElement('canvas');
    this.canvas.style.display = 'block';
    this.canvas.style.maxWidth = '100%';
    this.canvas.style.height = 'auto';
    container.appendChild(this.canvas);

    this.engine = new PptxPresentation({
      onReady: opts.onReady,
      onError: opts.onError,
    });
  }

  /** Load a PPTX file from an ArrayBuffer and render the first slide. */
  async load(buffer: ArrayBuffer): Promise<void> {
    await this.engine.load(buffer);
    this.currentSlide = 0;
    this.loadedSlideCount = this.engine.slideCount;
    await this.renderCurrentSlide();
  }

  /** Navigate to a specific slide (0-indexed). */
  async goToSlide(index: number): Promise<void> {
    if (this.loadedSlideCount === 0) return;
    this.currentSlide = Math.max(0, Math.min(index, this.loadedSlideCount - 1));
    await this.renderCurrentSlide();
  }

  async nextSlide(): Promise<void> {
    await this.goToSlide(this.currentSlide + 1);
  }

  async prevSlide(): Promise<void> {
    await this.goToSlide(this.currentSlide - 1);
  }

  get slideIndex(): number { return this.currentSlide; }
  get slideCount(): number { return this.loadedSlideCount; }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement { return this.canvas; }

  private async renderCurrentSlide(): Promise<void> {
    const targetWidth = this.opts.width ?? (this.canvas.offsetWidth || 960);
    const dpr = this.opts.dpr ?? (window.devicePixelRatio || 1);

    // Set CSS dimensions so the canvas displays at the correct aspect ratio
    const scale = targetWidth / this.engine.slideWidth;
    const cssHeight = Math.round(this.engine.slideHeight * scale);
    this.canvas.style.width = `${targetWidth}px`;
    this.canvas.style.height = `${cssHeight}px`;

    await this.engine.renderSlide(
      { canvas: this.canvas, width: targetWidth, dpr },
      this.currentSlide,
    );

    this.opts.onSlideChange?.(this.currentSlide, this.loadedSlideCount);
  }

  /** Clean up the viewer and terminate the background worker. */
  destroy(): void {
    this.engine.destroy();
    this.canvas.remove();
  }
}
