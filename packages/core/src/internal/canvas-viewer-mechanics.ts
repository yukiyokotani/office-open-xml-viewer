/**
 * Small browser-viewer mechanics shared by the format packages. This module is
 * intentionally an internal subpath: it owns DOM/canvas lifecycle facts, never
 * page/slide/sheet semantics or format state transitions.
 */

export type CanvasRestoreMode = 'display' | 'style-and-bitmap';

export const MAX_NATIVE_TEXT_SELECTION_CHARS = 65_536;
export const MAX_NATIVE_TEXT_SELECTION_LOCATORS = 1_024;

export interface BoundedNativeTextSelection<TLocator> {
  readonly text: string;
  readonly locators: readonly TLocator[];
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text' | 'runs')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
  readonly maxLocators: number;
}

function boundedSelectionLimit(value: number | undefined, maximum: number, name: string): number {
  const requested = value ?? maximum;
  if (!Number.isFinite(requested) || requested < 0) {
    throw new RangeError(`${name} must be a finite non-negative number.`);
  }
  return Math.min(maximum, Math.floor(requested));
}

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, maxCodeUnits);
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

function* descendantTextNodes(root: Node): Iterable<Text> {
  let node: Node | null = root.firstChild ?? root.childNodes[0] ?? null;
  while (node) {
    if (node.nodeType === 3) yield node as Text;
    if (node.firstChild) {
      node = node.firstChild;
      continue;
    }
    while (node && node !== root && !node.nextSibling) node = node.parentNode;
    if (!node || node === root) break;
    node = node.nextSibling;
  }
}

/** Append only the selected slice of one tagged run, never the full Selection. */
function appendSelectedRunText(
  chunks: string[],
  length: number,
  run: HTMLElement,
  ranges: readonly Range[],
  retainThrough: number,
): number {
  for (const range of ranges) {
    let intersects = false;
    try { intersects = range.intersectsNode(run); } catch { /* different DOM root */ }
    if (!intersects) continue;
    for (const node of descendantTextNodes(run)) {
      const value = node.data;
      let start: number | null;
      let end: number | null;
      if (range.startContainer === node) start = range.startOffset;
      else {
        try { start = range.comparePoint(node, 0) === 0 ? 0 : null; } catch { start = null; }
      }
      if (range.endContainer === node) end = range.endOffset;
      else {
        try { end = range.comparePoint(node, value.length) === 0 ? value.length : null; } catch { end = null; }
      }
      if (start === null || end === null || end <= start) continue;
      const available = retainThrough - length;
      if (available <= 0) return length;
      const chunk = value.slice(
        Math.max(0, start),
        Math.min(value.length, end, Math.max(0, start) + available),
      );
      chunks.push(chunk);
      length += chunk.length;
      if (length >= retainThrough) return length;
    }
  }
  return length;
}

/**
 * Read a browser-native text selection only when every range endpoint belongs
 * to this Viewer. This prevents a cross-DOM selection from leaking adjacent
 * page content into an AI/MCP context. Locators come only from tagged run spans
 * intersected by the native ranges and are detached by the caller's mapper.
 */
export function readBoundedNativeTextSelection<TLocator>(
  root: HTMLElement,
  selection: Selection | null,
  locatorForRun: (run: HTMLElement) => TLocator | null,
  options: Readonly<{ maxChars?: number; maxLocators?: number }> = {},
): BoundedNativeTextSelection<TLocator> | null {
  if (!selection || selection.isCollapsed || selection.rangeCount === 0) return null;
  const selectionSurfaces = [
    ...(root.matches?.('[data-ooxml-selection-surface]') ? [root] : []),
    ...root.querySelectorAll<HTMLElement>('[data-ooxml-selection-surface]'),
  ];
  if (selectionSurfaces.length === 0) return null;
  const isOnSelectionSurface = (node: Node) =>
    selectionSurfaces.some((surface) => surface.contains(node));
  const ranges: Range[] = [];
  for (let index = 0; index < selection.rangeCount; index++) {
    const range = selection.getRangeAt(index);
    if (!root.contains(range.startContainer) || !root.contains(range.endContainer) ||
        !isOnSelectionSurface(range.startContainer) ||
        !isOnSelectionSurface(range.endContainer)) return null;
    ranges.push(range);
  }
  const maxChars = boundedSelectionLimit(
    options.maxChars, MAX_NATIVE_TEXT_SELECTION_CHARS, 'maxTextCharacters',
  );
  const maxLocators = boundedSelectionLimit(
    options.maxLocators, MAX_NATIVE_TEXT_SELECTION_LOCATORS, 'maxRunLocators',
  );
  const locators: TLocator[] = [];
  let locatorOverflow = false;
  // Keep at most two look-ahead code units: one proves overflow and the second
  // lets safeUtf16Prefix observe a surrogate pair at the public boundary.
  const retainThrough = maxChars + 2;
  const textChunks: string[] = [];
  let retainedTextLength = 0;
  for (const candidate of root.querySelectorAll<HTMLElement>('[data-ooxml-selection-run]')) {
    const selected = ranges.some((range) => {
      try { return range.intersectsNode(candidate); } catch { return false; }
    });
    if (!selected) continue;
    const locator = locatorForRun(candidate);
    if (locator === null) continue;
    if (locators.length >= maxLocators) locatorOverflow = true;
    else locators.push(structuredClone(locator));
    if (retainedTextLength < retainThrough) {
      retainedTextLength = appendSelectedRunText(
        textChunks, retainedTextLength, candidate, ranges, retainThrough,
      );
    }
    if (locatorOverflow && retainedTextLength >= retainThrough) break;
  }
  if (locators.length === 0 && !locatorOverflow) return null;
  const retainedText = textChunks.join('');
  if (retainedText.length === 0) return null;
  const text = safeUtf16Prefix(retainedText, maxChars);
  const textOverflow = retainedText.length > maxChars;
  return {
    text,
    locators,
    truncated: textOverflow || locatorOverflow,
    truncationReasons: [
      ...(textOverflow ? ['text' as const] : []),
      ...(locatorOverflow ? ['runs' as const] : []),
    ],
    textCharacters: text.length,
    maxTextCharacters: maxChars,
    maxLocators,
  };
}

export interface CallerCanvasMountOptions {
  readonly wrapperCssText: string;
  readonly forceDisplayBlock?: boolean;
  readonly restoreMode?: CanvasRestoreMode;
}

/** Reparents a caller-owned canvas into one wrapper and restores it exactly. */
export class CallerCanvasMount {
  readonly wrapper: HTMLDivElement;

  private readonly originalParent: Node | null;
  private readonly originalNextSibling: Node | null;
  private readonly originalDisplay: string;
  private readonly originalStyle: string | null;
  private readonly originalWidth: number;
  private readonly originalHeight: number;
  private restored = false;

  constructor(
    readonly canvas: HTMLCanvasElement,
    private readonly options: CallerCanvasMountOptions,
  ) {
    this.originalParent = canvas.parentNode;
    this.originalNextSibling = canvas.nextSibling;
    this.originalDisplay = canvas.style.display;
    this.originalStyle = options.restoreMode === 'style-and-bitmap'
      ? canvas.getAttribute('style')
      : null;
    this.originalWidth = canvas.width;
    this.originalHeight = canvas.height;

    const ownerDocument = canvas.ownerDocument ?? document;
    this.wrapper = ownerDocument.createElement('div');
    this.wrapper.style.cssText = options.wrapperCssText;
    if (options.forceDisplayBlock && !canvas.style.display) canvas.style.display = 'block';
    if (this.originalParent) this.originalParent.insertBefore(this.wrapper, canvas);
    this.wrapper.appendChild(canvas);
  }

  /** Idempotently restore the original DOM slot and configured canvas state. */
  restore(): void {
    if (this.restored) return;
    this.restored = true;

    if (this.originalParent) {
      const reference = this.originalNextSibling?.parentNode === this.originalParent
        ? this.originalNextSibling
        : null;
      this.originalParent.insertBefore(this.canvas, reference);
    } else if (this.canvas.parentNode) {
      this.canvas.parentNode.removeChild(this.canvas);
    }

    if ((this.options.restoreMode ?? 'display') === 'style-and-bitmap') {
      if (this.originalStyle === null) this.canvas.removeAttribute('style');
      else this.canvas.setAttribute('style', this.originalStyle);
      this.canvas.width = this.originalWidth;
      this.canvas.height = this.originalHeight;
    } else {
      this.canvas.style.display = this.originalDisplay;
    }
    this.wrapper.remove();
  }
}

const TEXT_LAYER_STYLE =
  'position:absolute;top:0;left:0;width:100%;height:100%;' +
  'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
const HIGHLIGHT_LAYER_STYLE =
  'position:absolute;top:0;left:0;width:100%;height:100%;' +
  'overflow:hidden;pointer-events:none;';
const ELEMENT_OUTLINE_LAYER_STYLE =
  'position:absolute;top:0;left:0;width:100%;height:100%;' +
  'overflow:hidden;pointer-events:none;';

export interface CanvasElementOutline {
  /** Normalized coordinates relative to the page/slide canvas. */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation?: number;
}

/** Create the transparent layer used only by explicit element-context opt-in. */
export function createCanvasElementOutlineLayer(
  wrapper: HTMLElement,
  enabled: boolean,
): HTMLDivElement | null {
  if (!enabled) return null;
  const ownerDocument = wrapper.ownerDocument ?? document;
  const layer = ownerDocument.createElement('div');
  layer.style.cssText = ELEMENT_OUTLINE_LAYER_STYLE;
  wrapper.appendChild(layer);
  return layer;
}

/** Draw one non-editable selection frame without changing the document canvas. */
export function renderCanvasElementOutline(
  layer: HTMLDivElement | null,
  outline: CanvasElementOutline | null,
): void {
  if (!layer) return;
  layer.innerHTML = '';
  if (!outline || !Number.isFinite(outline.x) || !Number.isFinite(outline.y) ||
      !Number.isFinite(outline.width) || !Number.isFinite(outline.height) ||
      outline.width <= 0 || outline.height <= 0) return;
  const ownerDocument = layer.ownerDocument ?? document;
  const frame = ownerDocument.createElement('div');
  const rotation = Number.isFinite(outline.rotation) ? outline.rotation ?? 0 : 0;
  frame.style.cssText =
    `position:absolute;left:${outline.x * 100}%;top:${outline.y * 100}%;` +
    `width:${outline.width * 100}%;height:${outline.height * 100}%;` +
    'box-sizing:border-box;border:2px solid #1a73e8;' +
    'background:color-mix(in srgb, #1a73e8 6%, transparent);' +
    `transform:rotate(${rotation}deg);transform-origin:center;pointer-events:none;`;
  layer.appendChild(frame);
}

/** DOM containers only; each format remains responsible for overlay contents. */
export class CanvasOverlayHost {
  readonly textLayer: HTMLDivElement | null;
  readonly highlightLayer: HTMLDivElement;
  readonly elementLayer: HTMLDivElement | null;

  constructor(
    wrapper: HTMLElement,
    enableTextSelection: boolean,
    enableElementSelection = false,
  ) {
    const ownerDocument = wrapper.ownerDocument ?? document;
    this.textLayer = enableTextSelection ? ownerDocument.createElement('div') : null;
    if (this.textLayer) {
      this.textLayer.style.cssText = TEXT_LAYER_STYLE;
      wrapper.appendChild(this.textLayer);
    }
    this.highlightLayer = ownerDocument.createElement('div');
    this.highlightLayer.style.cssText = HIGHLIGHT_LAYER_STYLE;
    wrapper.appendChild(this.highlightLayer);
    this.elementLayer = createCanvasElementOutlineLayer(wrapper, enableElementSelection);
  }
}

export interface BitmapCommitSize {
  readonly cssWidth?: number;
  readonly cssHeight?: number;
}

export interface DestroyableResource {
  destroy(): void;
}

export type CanvasViewerRenderMode = 'main' | 'worker';

/** Resolve the mode of a viewer that may borrow an already-loaded engine. */
export function resolveCanvasViewerMode(
  viewerName: string,
  requestedMode: CanvasViewerRenderMode | undefined,
  engine: Readonly<{ mode: CanvasViewerRenderMode }> | undefined,
): CanvasViewerRenderMode {
  if (engine && requestedMode !== undefined && requestedMode !== engine.mode) {
    throw new Error(
      `${viewerName}: opts.mode='${requestedMode}' conflicts with the borrowed engine's ` +
        `mode='${engine.mode}'. Omit opts.mode when borrowing an engine — ` +
        'the engine owns its render mode.',
    );
  }
  return engine?.mode ?? requestedMode ?? 'main';
}

/**
 * Terminal, generation-safe ownership for a replaceable viewer resource.
 *
 * Page, slide, and sheet viewers share the same lifecycle invariant: a newer
 * acquisition supersedes an older one, a losing candidate is destroyed, and
 * close is permanent. Format-specific rendering remains outside this class.
 */
export class TerminalResourceOwner<T extends DestroyableResource> {
  private generation = 0;
  private resource: T | null;
  private ownsResource: boolean;
  private closed = false;

  constructor(
    private readonly ownerName: string,
    initial: T | null = null,
    ownsInitial = false,
  ) {
    this.resource = initial;
    this.ownsResource = initial !== null && ownsInitial;
  }

  get current(): T | null {
    return this.resource;
  }

  async replace(
    load: () => Promise<T>,
    beforeCommit?: (previous: T | null) => void,
  ): Promise<T | null> {
    this.assertOpen();
    const generation = ++this.generation;
    let candidate: T;
    try {
      candidate = await load();
    } catch (error) {
      if (this.closed) throw this.closedError();
      if (generation !== this.generation) return null;
      throw error;
    }
    if (this.closed) {
      this.dispose(candidate);
      throw this.closedError();
    }
    if (generation !== this.generation) {
      this.dispose(candidate);
      return null;
    }
    try {
      beforeCommit?.(this.resource);
    } catch (error) {
      this.dispose(candidate);
      throw error;
    }
    this.install(candidate, true);
    return candidate;
  }

  install(candidate: T, owned = true): void {
    this.assertOpen();
    // A direct installation is itself a replacement generation. Any loader
    // already in flight must lose when it resolves; otherwise it can overwrite
    // this explicitly installed resource and destroy it as the previous owner.
    this.generation++;
    const previous = this.resource;
    const ownedPrevious = this.ownsResource;
    this.resource = candidate;
    this.ownsResource = owned;
    if (ownedPrevious && previous) this.dispose(previous);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.generation++;
    const previous = this.resource;
    const ownedPrevious = this.ownsResource;
    this.resource = null;
    this.ownsResource = false;
    if (ownedPrevious && previous) this.dispose(previous);
  }

  private assertOpen(): void {
    if (this.closed) throw this.closedError();
  }

  private closedError(): Error {
    return new Error(`${this.ownerName} is closed`);
  }

  /** Cleanup cannot change the already-committed ownership transition. */
  private dispose(resource: T): void {
    try { resource.destroy(); } catch {}
  }
}

/**
 * Generation gate for a single static canvas. Format renderers still perform
 * their own drawing; this object prevents stale completion side effects and
 * owns worker ImageBitmap commit/disposal.
 */
export class StaticCanvasRenderDispatcher {
  private generation = 0;
  private destroyed = false;
  private readonly bitmapContext: ImageBitmapRenderingContext | null;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    bitmapMode: boolean,
  ) {
    this.bitmapContext = bitmapMode ? canvas.getContext('bitmaprenderer') : null;
  }

  begin(): number {
    return ++this.generation;
  }

  isCurrent(generation: number): boolean {
    return !this.destroyed && generation === this.generation;
  }

  /** Commit a worker bitmap only if it still belongs to the active generation. */
  commitBitmap(
    generation: number,
    bitmap: ImageBitmap,
    size: BitmapCommitSize = {},
  ): boolean {
    if (!this.isCurrent(generation)) {
      bitmap.close();
      return false;
    }
    if (!this.bitmapContext) {
      bitmap.close();
      throw new Error('bitmaprenderer context not available');
    }
    if (this.canvas.width !== bitmap.width) this.canvas.width = bitmap.width;
    if (this.canvas.height !== bitmap.height) this.canvas.height = bitmap.height;
    if (size.cssWidth !== undefined) this.canvas.style.width = `${size.cssWidth}px`;
    if (size.cssHeight !== undefined) this.canvas.style.height = `${size.cssHeight}px`;
    try {
      this.bitmapContext.transferFromImageBitmap(bitmap);
    } catch (error) {
      bitmap.close();
      throw error;
    }
    return true;
  }

  /**
   * Commit a worker bitmap through a 2D context. This keeps ImageBitmap
   * ownership and stale-generation disposal inside the dispatcher while
   * allowing canvases that must remain compatible with interactive 2D media
   * rendering to avoid acquiring a `bitmaprenderer` context.
   */
  commitBitmapTo2d(
    generation: number,
    bitmap: ImageBitmap,
    size: BitmapCommitSize = {},
  ): boolean {
    if (!this.isCurrent(generation)) {
      bitmap.close();
      return false;
    }
    if (this.canvas.width !== bitmap.width) this.canvas.width = bitmap.width;
    if (this.canvas.height !== bitmap.height) this.canvas.height = bitmap.height;
    if (size.cssWidth !== undefined) this.canvas.style.width = `${size.cssWidth}px`;
    if (size.cssHeight !== undefined) this.canvas.style.height = `${size.cssHeight}px`;
    const context = this.canvas.getContext('2d');
    if (!context) {
      bitmap.close();
      throw new Error('2D context not available');
    }
    try {
      context.drawImage(bitmap, 0, 0);
    } finally {
      bitmap.close();
    }
    return true;
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.generation++;
  }
}

/** Shared render-error delivery, with a permanent close gate for teardown. */
export class CanvasViewerErrorRouter {
  private closed = false;
  private readonly handled = new WeakSet<Error>();
  /** Awaiters for the same background lifecycle reported by reportBackground.
   * This is deliberately narrower than all Viewer promises: an unrelated find
   * or render operation must never suppress a terminal layout notification. */
  private backgroundLifecycleOwners = 0;

  constructor(
    private readonly viewerName: string,
    private readonly onError?: (error: Error) => void,
  ) {}

  report(error: unknown): void {
    if (this.closed) return;
    const normalized = error instanceof Error ? error : new Error(String(error));
    if (this.handled.has(normalized)) return;
    this.handled.add(normalized);
    if (this.onError) this.onError(normalized);
    else console.error(`[ooxml] ${this.viewerName} render failed:`, normalized);
  }

  /** Claim an Error for another explicit callback so derived render rejections
   * cannot deliver the same terminal failure again through onError. */
  markHandled(error: unknown): void {
    if (this.closed || !(error instanceof Error)) return;
    this.handled.add(error);
  }

  /** Keep a terminal background failure on the Promise channel while an
   * explicit awaitable Viewer operation is waiting for that same lifecycle.
   * Mark the exact Error identity before rethrowing; unrelated background
   * failures must never be hidden merely because another Promise is pending. */
  async ownAwaitable<T>(operation: () => Promise<T>): Promise<T> {
    try {
      return await operation();
    } catch (error) {
      this.markHandled(error);
      throw error;
    }
  }

  /** Claim the background lifecycle while an explicit public Promise waits for
   * it. Publication and Promise rejection can cross any number of async jobs;
   * synchronous ownership makes their single-channel contract deterministic. */
  async ownBackgroundLifecycle<T>(operation: () => Promise<T>): Promise<T> {
    this.backgroundLifecycleOwners++;
    try {
      return await this.ownAwaitable(operation);
    } finally {
      this.backgroundLifecycleOwners--;
    }
  }

  /** Route an unowned background failure, or claim it for an explicit callback
   * owner or an explicit Promise awaiting this same lifecycle. */
  reportBackground(error: unknown, explicitCallbackOwner = false): void {
    const normalized = error instanceof Error ? error : new Error(String(error));
    if (explicitCallbackOwner || this.backgroundLifecycleOwners > 0) {
      this.markHandled(normalized);
      return;
    }
    this.report(normalized);
  }

  close(): void {
    this.closed = true;
  }
}
