import {
  dropDecodedBitmapCache,
  dropSvgImageCache,
  OoxmlResourceLimitError,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import type { Presentation, Slide } from '@silurus/ooxml-pptx';
import {
  decodeOoxmlResourceUsage,
  parseResourceLimitError,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  type PullSessionCommand,
  type PullSessionResponse,
  type OoxmlResourceMetricsSession,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { usingOwnedSession } from '@silurus/ooxml-core/internal/owned-session';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
import {
  acquirePptxNodeSession,
  PptxSlidePullClient,
  readPptxSlideCursorUsage,
  SlidePullWorker,
  type PresentationBootstrap,
  type PptxNodeArchive,
} from '@silurus/ooxml-pptx/internal/session';
import { InProcessPullTransport } from '@silurus/ooxml-core/internal/in-process-pull-transport';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import type { NodeCanvasFactory, NodeCanvasLike } from './render.ts';
import { createLazyWasmModule, resolveWasm } from './wasm-loader.ts';
import { normalizeNodeOfficeInput } from './normalize-input.ts';

const getPptxWasmModule = createLazyWasmModule(() => resolveWasm(
    import.meta.url,
    'pptx_parser_bg.wasm',
    '@silurus/ooxml-pptx/wasm-binary',
  ));

/** Options for the bounded Node presentation session. */
export type OpenPptxPresentationOptions = OoxmlNodeSessionOptions;

export interface PptxSessionRenderOptions {
  readonly width?: number;
  readonly dpr?: number;
  /** Required to decode images and allocate effect surfaces under Node. */
  readonly factory: NodeCanvasFactory;
}

/**
 * One-pass Node session over canonical complete PPTX slide units. The library
 * retains at most the package bootstrap and one yielded slide at a time. Slides
 * deliberately remain ordinary objects: copies retained by the caller are
 * caller-owned and are outside this session's memory bound.
 */
export interface PptxPresentationSession extends AsyncIterable<Slide> {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  getImage(path: string, mimeType: string): Promise<Blob>;
  getMedia(path: string, mimeType?: string): Promise<Blob>;
  renderSlide(
    canvas: NodeCanvasLike,
    slide: Slide,
    options: PptxSessionRenderOptions,
  ): Promise<void>;
  slides(): AsyncGenerator<Slide, void, void>;
  close(): Promise<void>;
}

/**
 * Open a one-pass, pull-based PPTX session for Node batch rendering. Existing
 * materializing helpers remain unchanged; this additive path avoids retaining a
 * complete `Presentation` while callers render or otherwise consume each slide.
 * Exhausting or breaking the iterator closes the retained WASM archive.
 */
export async function openPptxPresentation(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenPptxPresentationOptions = {},
): Promise<PptxPresentationSession> {
  return openPptxPresentationImpl(buffer, options);
}

async function openPptxPresentationImpl(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenPptxPresentationOptions = {},
): Promise<PptxPresentationSessionImpl> {
  const bytes = await normalizeNodeOfficeInput(buffer, 'pptx', options);
  const acquired = await acquirePptxNodeSession(bytes, getPptxWasmModule(), options);
  return new PptxPresentationSessionImpl(
    acquired.closeArchive,
    acquired.archive,
    acquired.bootstrap,
    acquired.metrics,
    options.signal,
  );
}

class PptxPresentationSessionImpl implements PptxPresentationSession {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;

  private readonly slidePull: SlidePullWorker;
  private readonly slideClient: PptxSlidePullClient;
  private readonly transport: InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>>;
  private started = false;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private usage: OoxmlResourceUsageSnapshot | undefined;
  private consumedSlides = 0;
  private resourceFailure: OoxmlResourceLimitError | undefined;
  private renderTail: Promise<void> = Promise.resolve();
  private readonly fetchImage = (path: string, mimeType: string): Promise<Blob> =>
    this.getPartInternal(path, mimeType, (archive) => archive.extract_image(path));
  private readonly fetchMedia = (path: string): Promise<Blob> =>
    this.getPartInternal(path, 'application/octet-stream', (archive) => archive.extract_media(path));
  private readonly rawParts = new BoundedRawPartCache({
    maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
    maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
  });

  constructor(
    private readonly closeArchive: () => void,
    private readonly archive: PptxNodeArchive,
    private readonly bootstrap: PresentationBootstrap,
    private readonly metrics: OoxmlResourceMetricsSession,
    private readonly signal?: AbortSignal,
  ) {
    this.slideCount = bootstrap.slideCount;
    this.slideWidth = bootstrap.slideWidth;
    this.slideHeight = bootstrap.slideHeight;
    this.slidePull = new SlidePullWorker(() => this.archive);
    this.transport = new InProcessPullTransport(
      (command, respond) => this.slidePull.dispatchSafely(
        command as PullSessionCommand<number>,
        respond,
      ),
      () => undefined,
    );
    this.slideClient = new PptxSlidePullClient({
      slideCount: this.slideCount,
      transport: this.transport,
      open: async (slideIndex, identity) => {
        this.slidePull.reserveOpen(identity);
        await this.slidePull.open(slideIndex, identity);
      },
      onUsage: (usage) => {
        this.usage = usage;
        this.metrics.observeUsage(usage);
      },
    });
  }

  materialize(slides: Slide[]): Presentation {
    return {
      slideWidth: this.slideWidth,
      slideHeight: this.slideHeight,
      slides,
      defaultTextColor: this.bootstrap.defaultTextColor,
      majorFont: this.bootstrap.majorFont,
      minorFont: this.bootstrap.minorFont,
      ...(this.bootstrap.hlinkColor ? { hlinkColor: this.bootstrap.hlinkColor } : {}),
      ...(this.bootstrap.folHlinkColor ? { folHlinkColor: this.bootstrap.folHlinkColor } : {}),
    };
  }

  get resourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    if (this.closed) return this.usage;
    return this.refreshResourceUsage();
  }

  async getImage(path: string, mimeType: string): Promise<Blob> {
    this.assertOpen();
    return this.getPartInternal(path, mimeType, (archive) => archive.extract_image(path))
      .catch((error: unknown) => this.failOperation(error));
  }

  async getMedia(path: string, mimeType = 'application/octet-stream'): Promise<Blob> {
    this.assertOpen();
    return this.getPartInternal(path, mimeType, (archive) => archive.extract_media(path))
      .catch((error: unknown) => this.failOperation(error));
  }

  async renderSlide(
    canvas: NodeCanvasLike,
    slide: Slide,
    options: PptxSessionRenderOptions,
  ): Promise<void> {
    this.assertOpen();
    return this.enqueueRender(async () => {
      throwIfAborted(this.signal);
      const { renderSlideNode } = await import('./render.ts');
      const presentation: Presentation = {
        slideWidth: this.slideWidth,
        slideHeight: this.slideHeight,
        slides: [slide],
        defaultTextColor: this.bootstrap.defaultTextColor,
        majorFont: this.bootstrap.majorFont,
        minorFont: this.bootstrap.minorFont,
        ...(this.bootstrap.hlinkColor ? { hlinkColor: this.bootstrap.hlinkColor } : {}),
        ...(this.bootstrap.folHlinkColor ? { folHlinkColor: this.bootstrap.folHlinkColor } : {}),
      };
      await renderSlideNode(canvas, presentation, 0, {
        ...options,
        fetchImage: this.fetchImage,
        fetchMedia: this.fetchMedia,
      });
      throwIfAborted(this.signal);
    }).catch((error: unknown) => this.failOperation(error));
  }

  [Symbol.asyncIterator](): AsyncGenerator<Slide, void, void> {
    return this.slides();
  }

  async *slides(): AsyncGenerator<Slide, void, void> {
    if (this.closed) throw new Error('PPTX presentation session is closed');
    if (this.started) throw new Error('PPTX presentation session is one-pass and was already consumed');
    this.started = true;
    let operationError: unknown;
    try {
      for (let index = 0; index < this.slideCount; index += 1) {
        throwIfAborted(this.signal);
        const slide = await this.slideClient.load(index);
        if (!slide) throw new Error(`PPTX slide ${index} was not decoded`);
        this.usage ??= await this.slidePull.run(() => readPptxSlideCursorUsage(
          (operation) => operation(this.archive),
        ));
        this.metrics.observeUsage(this.usage);
        this.consumedSlides = index + 1;
        yield slide;
      }
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
      this.metrics.fail(operationError);
      throw operationError;
    } finally {
      try {
        await this.close();
      } catch (cleanupError) {
        if (operationError === undefined) throw cleanupError;
      }
    }
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closed = true;
    this.slideClient.cancelAll();
    this.closePromise = this.release();
    return this.closePromise;
  }

  private async release(): Promise<void> {
    let operationError: unknown;
    await this.renderTail;
    dropDecodedBitmapCache(this.fetchImage);
    dropSvgImageCache(this.fetchImage);
    try {
      await this.slidePull.reset();
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
    }
    this.transport.terminate();
    this.rawParts.clear();
    try {
      this.closeArchive();
    } catch (cleanupError) {
      operationError ??= parseResourceLimitError(cleanupError) ?? cleanupError;
    }
    if (operationError !== undefined) {
      this.metrics.fail(operationError);
      throw operationError;
    }
    this.metrics.checkpoint('presentation session closed');
    this.metrics.succeed({ slides: this.consumedSlides });
  }

  private enqueueRender<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.renderTail.then(operation, operation);
    this.renderTail = result.then(() => undefined, () => undefined);
    return result;
  }

  private getPartInternal(
    path: string,
    mimeType: string,
    extract: (archive: PptxNodeArchive) => Uint8Array,
  ): Promise<Blob> {
    return this.rawParts.get(path, mimeType, () => {
      throwIfAborted(this.signal);
      const bytes = extract(this.archive);
      this.refreshResourceUsage();
      return new Blob([bytes as BlobPart], { type: mimeType });
    });
  }

  private refreshResourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    try {
      this.usage = decodeOoxmlResourceUsage(this.archive.resource_usage());
      this.metrics.observeUsage(this.usage);
    } catch {
      // Preserve the last complete diagnostic if the archive is closing/trapped.
    }
    return this.usage;
  }

  private assertOpen(): void {
    if (this.closed) throw new Error('PPTX presentation session is closed');
    if (this.resourceFailure) throw this.resourceFailure;
  }

  private failOperation(error: unknown): never {
    const normalized = parseResourceLimitError(error) ?? error;
    if (normalized instanceof OoxmlResourceLimitError) {
      this.resourceFailure ??= normalized;
    }
    this.metrics.fail(normalized);
    throw normalized;
  }
}

/** Materialize a complete caller-owned presentation through the same retained
 * archive and acknowledged slide producer as {@link openPptxPresentation}. */
export async function materializePptxPresentation(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenPptxPresentationOptions = {},
): Promise<Presentation> {
  return usingOwnedSession(
    () => openPptxPresentationImpl(buffer, options),
    async (session) => {
      const slides: Slide[] = [];
      for await (const slide of session.slides()) slides.push(slide);
      return session.materialize(slides);
    },
  );
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('PPTX presentation session was aborted');
  error.name = 'AbortError';
  throw error;
}
