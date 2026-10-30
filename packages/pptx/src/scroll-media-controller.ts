import { StaticCanvasRenderDispatcher } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { VisibleWindow } from '@silurus/ooxml-core/internal/virtual-scroll';
import type { HyperlinkTarget } from '@silurus/ooxml-core';
import type { PptxPresentation } from './presentation';
import type { PresentationHandle } from './presentation-handle';
import type { PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';

export interface MediaSlot {
  canvas: HTMLCanvasElement;
  wrapper: HTMLDivElement;
  textLayer: HTMLDivElement | null;
  dispatcher: StaticCanvasRenderDispatcher;
  renderedSlide: number;
  renderedScale: number;
  mediaInteractive: boolean;
  presentationGeneration: number;
  presentationHandle: PresentationHandle | null;
}

interface MediaHooks<Slot extends MediaSlot> {
  presentation(): PptxPresentation | null;
  slots(): ReadonlyMap<number, Slot>;
  epoch(): number;
  scale(): number;
  dpr(): number;
  width(): number;
  height(): number;
  imageResources(): NonNullable<Parameters<PptxPresentation['presentSlide']>[2]>['imageResources'];
  errorPlaceholder?(): NonNullable<Parameters<PptxPresentation['presentSlide']>[2]>['errorPlaceholder'];
  textSelection(): boolean;
  findActive(): boolean;
  mediaEnabled(): boolean;
  mediaRange(): VisibleWindow;
  rangeContains(range: VisibleWindow, slide: number): boolean;
  shadow(): string | false;
  hyperlinkHandler(): ((target: HyperlinkTarget) => void) | undefined;
  refreshFindRuns(slide: number, runs: PptxTextRunInfo[]): void;
  redrawHighlights(slide: number, slot: Slot): void;
  commitComments(slide: number, slot: Slot): void;
  clearTextPreview(layer: HTMLDivElement): void;
  reportError(error: unknown): void;
}

/** PPTX-only interactive video presentation lifecycle, separate from the
 * shared static slot renderers. Handles belong to the canvas they painted. */
export class PptxScrollMediaController<Slot extends MediaSlot> {
  constructor(private readonly hooks: MediaHooks<Slot>) {}

  renderInteractive(i: number, slot: Slot, width: number, dpr: number,
    scale: number, epoch: number, reportErrors = true): Promise<void> {
    const presentation = this.hooks.presentation();
    if (!presentation) return Promise.resolve();
    const generation = ++slot.presentationGeneration;
    slot.presentationHandle?.destroy();
    slot.presentationHandle = null;
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = this.hooks.textSelection() && !!slot.textLayer;
    const wantRuns = wantOverlay || this.hooks.findActive();
    return presentation.presentSlide(slot.canvas, i, {
      width, dpr, imageResources: this.hooks.imageResources(),
      ...this.errorPlaceholderOption(),
      onTextRun: wantRuns ? (run) => runs.push(run) : undefined,
      onError: (error) => {
        if (generation === slot.presentationGeneration) this.hooks.reportError(error);
      },
    }).then((handle) => {
      if (!this.isCurrent(i, slot, generation, epoch)) {
        handle.destroy();
        return;
      }
      slot.presentationHandle = handle;
      slot.renderedScale = scale;
      if (wantOverlay && slot.textLayer) {
        buildPptxTextLayer(slot.textLayer, runs, Math.round(width),
          Math.round(this.hooks.height()), this.hooks.hyperlinkHandler(), i);
      }
      if (wantRuns) this.hooks.refreshFindRuns(i, runs);
      this.hooks.commitComments(i, slot);
      this.hooks.redrawHighlights(i, slot);
    }).catch((error: unknown) => {
      if (generation !== slot.presentationGeneration) return;
      if (reportErrors) this.hooks.reportError(error);
      else throw error;
    });
  }

  sync(range = this.hooks.mediaRange()): void {
    if (!this.hooks.mediaEnabled()) return;
    for (const [i, slot] of this.hooks.slots()) {
      const interactive = this.hooks.rangeContains(range, i);
      if (interactive === slot.mediaInteractive) continue;
      if (interactive) {
        slot.mediaInteractive = true;
        this.settleInteractive(i, slot, this.hooks.width(), this.hooks.dpr(),
          this.hooks.scale(), this.hooks.epoch());
      } else {
        slot.mediaInteractive = false;
        slot.presentationGeneration++;
        slot.presentationHandle?.destroy();
        slot.presentationHandle = null;
      }
    }
  }

  settleInteractive(i: number, slot: Slot, width: number, dpr: number,
    scale: number, epoch: number): void {
    const presentation = this.hooks.presentation();
    if (!presentation) return;
    const generation = ++slot.presentationGeneration;
    const spare = document.createElement('canvas');
    spare.style.cssText = 'display:block;background:#fff;';
    const shadow = this.hooks.shadow();
    if (shadow !== false) spare.style.boxShadow = shadow;
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = this.hooks.textSelection() && !!slot.textLayer;
    const wantRuns = wantOverlay || this.hooks.findActive();
    void presentation.presentSlide(spare, i, {
      width, dpr,
      ...this.errorPlaceholderOption(),
      onTextRun: wantRuns ? (run) => runs.push(run) : undefined,
      onError: (error) => {
        if (generation === slot.presentationGeneration) this.hooks.reportError(error);
      },
    }).then((handle) => {
      if (!this.isCurrent(i, slot, generation, epoch)) {
        handle.destroy();
        return;
      }
      const oldCanvas = slot.canvas;
      const oldHandle = slot.presentationHandle;
      slot.dispatcher.destroy();
      slot.wrapper.insertBefore(spare, oldCanvas);
      oldCanvas.remove();
      slot.canvas = spare;
      slot.dispatcher = new StaticCanvasRenderDispatcher(spare, false);
      slot.presentationHandle = handle;
      slot.renderedScale = scale;
      oldHandle?.destroy();
      if (slot.textLayer) {
        this.hooks.clearTextPreview(slot.textLayer);
        if (wantOverlay) {
          buildPptxTextLayer(slot.textLayer, runs, Math.round(width),
            Math.round(this.hooks.height()), this.hooks.hyperlinkHandler(), i);
        }
      }
      if (wantRuns) this.hooks.refreshFindRuns(i, runs);
      this.hooks.commitComments(i, slot);
      this.hooks.redrawHighlights(i, slot);
    }).catch((error: unknown) => {
      if (generation === slot.presentationGeneration) this.hooks.reportError(error);
    });
  }

  private isCurrent(i: number, slot: Slot, generation: number, epoch: number): boolean {
    return generation === slot.presentationGeneration && slot.mediaInteractive &&
      epoch === this.hooks.epoch() && this.hooks.slots().get(i) === slot &&
      slot.renderedSlide === i;
  }

  private errorPlaceholderOption() {
    const errorPlaceholder = this.hooks.errorPlaceholder?.();
    return errorPlaceholder === undefined ? {} : { errorPlaceholder };
  }

  destroy(): void {
    for (const slot of this.hooks.slots().values()) {
      slot.presentationGeneration++;
      slot.presentationHandle?.destroy();
      slot.presentationHandle = null;
      slot.mediaInteractive = false;
    }
  }
}
