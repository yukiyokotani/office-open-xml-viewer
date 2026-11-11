import { CanvasOverlayHost } from '../internal/canvas-viewer-mechanics';

/** UI gesture debounce; repeated CSS previews settle into one full render. */
export const DEFAULT_ZOOM_SETTLE_MS = 150;

/** The geometry can use either per-unit offsets or a uniform stride. Only the
 * visible window is needed by the slot pool. */
export interface SlotWindow {
  start: number;
  end: number;
  topIndex: number;
  totalHeight: number;
}

export interface SlotScrollerHooks<Slot, Range extends SlotWindow> {
  spacer(): HTMLDivElement;
  count(): number;
  range(): Range;
  createSlot(): Slot;
  attachSlot(slot: Slot): void;
  resetSlot(index: number, slot: Slot): void;
  positionSlot(index: number, slot: Slot, range: Range): void;
  renderSlot(index: number, slot: Slot, reportErrors: boolean): Promise<void> | null | void;
  previewSlot(index: number, slot: Slot, range: Range): void;
  settleSlot(index: number, slot: Slot): void;
  renderedScale(slot: Slot): number;
  scale(): number;
  syncSpacerWidth(): void;
  onRange(range: Range): void;
  onNewSlot?(index: number, slot: Slot): void;
  onExistingSlot?(index: number, slot: Slot, reportErrors: boolean): Promise<void> | null | void;
  afterMount?(range: Range): void;
  awaitInitialRender?(index: number): boolean;
  shouldSettle?(index: number, slot: Slot): boolean;
  /** Maps a layout position of {@link range} to the unit index that owns it.
   *  Omitted: identity. A viewer that leaves units out of its layout (PPTX
   *  hidden slides in `'skip'` mode) lays out positions and keys slots by unit. */
  unitAt?(position: number): number;
}

/** Owns bounded slot mounting, reuse and preview/settle scheduling. Render and
 * unit geometry remain hooks because DOCX pages vary while PPTX slides are
 * uniform and can own interactive media handles. */
export class SlotScroller<Slot, Range extends SlotWindow> {
  readonly slots = new Map<number, Slot>();
  readonly free: Slot[] = [];
  readonly inFlight = new Set<number>();
  lastRange: Range | null = null;
  renderEpoch = 0;
  private _settleTimer: ReturnType<typeof setTimeout> | null = null;
  private _destroyed = false;

  constructor(private readonly hooks: SlotScrollerHooks<Slot, Range>) {}

  syncSpacer(): void {
    const range = this.hooks.range();
    this.lastRange = range;
    this.hooks.spacer().style.height = `${range.totalHeight}px`;
    this.hooks.syncSpacerWidth();
  }

  acquireSlot(): Slot {
    const reused = this.free.pop();
    if (reused) {
      this.hooks.attachSlot(reused);
      return reused;
    }
    return this.hooks.createSlot();
  }

  recycleSlot(index: number, slot: Slot): void {
    this.slots.delete(index);
    this.hooks.resetSlot(index, slot);
    this.free.push(slot);
  }

  mount(initialRenders?: Promise<void>[], repositionExisting = true): void {
    if (this._destroyed || this.hooks.count() === 0) return;
    const range = this.hooks.range();
    this.lastRange = range;
    const units = this.unitsIn(range);
    for (const [index, slot] of [...this.slots]) {
      if (!units.has(index)) this.recycleSlot(index, slot);
    }
    for (const index of units) {
      const existing = this.slots.get(index);
      if (existing) {
        if (!repositionExisting) continue;
        this.hooks.positionSlot(index, existing, range);
        const render = this.hooks.onExistingSlot?.(index, existing, initialRenders === undefined);
        if (initialRenders && render) initialRenders.push(render);
        continue;
      }
      const slot = this.acquireSlot();
      this.hooks.positionSlot(index, slot, range);
      this.slots.set(index, slot);
      this.hooks.onNewSlot?.(index, slot);
      const render = this.hooks.renderSlot(index, slot, initialRenders === undefined);
      if (initialRenders && render && (this.hooks.awaitInitialRender?.(index) ?? true)) {
        initialRenders.push(render);
      }
    }
    this.hooks.afterMount?.(range);
    this.hooks.onRange(range);
  }

  preview(): void {
    // Slots already mounted keep their device buffers; only CSS geometry grows
    // until a debounced settle paints a crisp replacement. New entrants paint
    // at the current scale directly, so they have nothing to stretch.
    if (this._destroyed || this.hooks.count() === 0) return;
    const range = this.hooks.range();
    this.lastRange = range;
    const units = this.unitsIn(range);
    for (const [index, slot] of [...this.slots]) {
      if (!units.has(index)) this.recycleSlot(index, slot);
    }
    for (const index of units) {
      const existing = this.slots.get(index);
      if (existing) {
        this.hooks.previewSlot(index, existing, range);
        continue;
      }
      const slot = this.acquireSlot();
      this.hooks.positionSlot(index, slot, range);
      this.slots.set(index, slot);
      this.hooks.onNewSlot?.(index, slot);
      this.hooks.renderSlot(index, slot, true);
    }
    this.hooks.afterMount?.(range);
    this.hooks.onRange(range);
  }

  /** Units owning the window's positions, in layout order. */
  private unitsIn(range: Range): Set<number> {
    const units = new Set<number>();
    const unitAt = this.hooks.unitAt;
    for (let position = range.start; position <= range.end; position++) {
      units.add(unitAt ? unitAt(position) : position);
    }
    return units;
  }

  scheduleSettle(delayMs: number): void {
    if (this._settleTimer !== null) clearTimeout(this._settleTimer);
    this._settleTimer = setTimeout(() => {
      this._settleTimer = null;
      this.settle();
    }, delayMs);
  }

  settle(): void {
    if (this._destroyed || this.hooks.count() === 0) return;
    for (const [index, slot] of [...this.slots]) {
      if (this.hooks.renderedScale(slot) === this.hooks.scale()) continue;
      if (this.hooks.shouldSettle && !this.hooks.shouldSettle(index, slot)) continue;
      this.hooks.settleSlot(index, slot);
    }
  }

  destroy(): void {
    if (this._destroyed) return;
    this._destroyed = true;
    if (this._settleTimer !== null) clearTimeout(this._settleTimer);
    this._settleTimer = null;
    for (const [index, slot] of [...this.slots]) this.recycleSlot(index, slot);
    this.inFlight.clear();
    this.free.length = 0;
  }
}

/** Restore an overlay after transient CSS zoom preview. */
export function clearTextLayerPreview(layer: HTMLDivElement): void {
  layer.style.transform = '';
  layer.style.transformOrigin = '';
  layer.style.width = '100%';
  layer.style.height = '100%';
}

export interface SlotHost {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  highlightLayer: HTMLDivElement;
}

/** The common canvas, selection and highlight stack. Formats append media,
 * comment and element layers in their own paint order. */
export function createSlotHost(
  scrollHost: HTMLDivElement,
  textSelection: boolean,
  shadow: string | false,
): SlotHost {
  const doc = scrollHost.ownerDocument;
  const wrapper = doc.createElement('div');
  wrapper.style.cssText = 'position:absolute;';
  const canvas = doc.createElement('canvas');
  canvas.style.cssText = 'display:block;background:#fff;';
  if (shadow !== false) canvas.style.boxShadow = shadow;
  wrapper.appendChild(canvas);
  const { textLayer, highlightLayer } = new CanvasOverlayHost(wrapper, textSelection);
  return { wrapper, canvas, textLayer, highlightLayer };
}

/** Apply bitmap and text-overlay CSS preview without touching the backing
 * canvas buffer. Returns the zoom ratio for format-specific comment layers. */
export function previewSlotHost(
  slot: SlotHost & { renderedScale: number },
  width: number,
  height: number,
  scale: number,
): number | null {
  slot.canvas.style.width = `${width}px`;
  slot.canvas.style.height = `${height}px`;
  if (slot.renderedScale <= 0) return null;
  const ratio = scale / slot.renderedScale;
  if (slot.textLayer) {
    slot.textLayer.style.transformOrigin = '0 0';
    // Keep the box at committed dimensions: scaling a 100% box after its
    // wrapper grows would inflate the native scroll extent a second time.
    slot.textLayer.style.width = `${width / ratio}px`;
    slot.textLayer.style.height = `${height / ratio}px`;
    slot.textLayer.style.transform = `scale(${ratio})`;
  }
  return ratio;
}

export function resetSlotHost(slot: SlotHost): void {
  if (slot.textLayer) {
    slot.textLayer.innerHTML = '';
    clearTextLayerPreview(slot.textLayer);
  }
  slot.highlightLayer.innerHTML = '';
  slot.highlightLayer.style.transform = '';
  slot.highlightLayer.style.transformOrigin = '';
}

export interface CommentSlotLayers {
  markerLayer: HTMLDivElement | null;
  margin: HTMLDivElement | null;
  decorationLayer: HTMLDivElement | null;
}

/** Mount the format-neutral comment surfaces. The adapter computes margin
 * geometry and supplies the format's comment content and event handlers. */
export function createCommentSlotLayers(
  wrapper: HTMLDivElement,
  enabled: boolean,
  cards: boolean,
  connectors: boolean,
  syncMargin: (margin: HTMLDivElement) => void,
): CommentSlotLayers {
  if (!enabled) return { markerLayer: null, margin: null, decorationLayer: null };
  const doc = wrapper.ownerDocument;
  const markerLayer = doc.createElement('div');
  markerLayer.setAttribute('data-overlay', 'comment-marker');
  markerLayer.style.cssText = 'position:absolute;inset:0;overflow:hidden;pointer-events:none;';
  wrapper.appendChild(markerLayer);
  if (!cards) return { markerLayer, margin: null, decorationLayer: null };
  const margin = doc.createElement('div');
  margin.style.cssText =
    'position:absolute;top:0;height:100%;box-sizing:border-box;' +
    'overflow-x:hidden;overflow-y:auto;pointer-events:auto;';
  syncMargin(margin);
  let decorationLayer: HTMLDivElement | null = null;
  if (connectors) {
    decorationLayer = doc.createElement('div');
    decorationLayer.style.cssText = 'position:absolute;top:0;left:0;overflow:visible;pointer-events:none;';
    wrapper.appendChild(decorationLayer);
  }
  wrapper.appendChild(margin);
  return { markerLayer, margin, decorationLayer };
}
