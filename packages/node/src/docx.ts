import { resolveCjkFallback, type CjkFallback, type ModelSource } from '@silurus/ooxml-core';
import type { DocxDocumentModel, DocxTextRunInfo } from '@silurus/ooxml-docx';
import type {
  OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  dropDecodedBitmapCache,
  dropSvgImageCache,
  OoxmlResourceLimitError,
} from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  parseResourceLimitError,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  type OoxmlResourceMetricsSession,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import {
  acquireDocxNodeDocument,
  normalizeDocxDocumentModel,
  materializeDocumentPullLayoutSession,
  materializeDocumentPullSession,
  type DocxNodeArchive,
  createLayoutServices,
  retainRenderWorkerDocumentLayout,
  renderLayoutSourceToCanvas,
} from '@silurus/ooxml-docx/internal/session';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import { normalizeCurrentDate, toUint8 } from './input-normalization.ts';
import {
  type NodeCanvasFactory,
  type NodeCanvasLike,
  withNodeCanvasRuntime,
} from './render.ts';
import { createLazyWasmModule, resolveWasm } from './wasm-loader.ts';
import { usingOwnedSession } from '@silurus/ooxml-core/internal/owned-session';
import { throwIfAborted } from '@silurus/ooxml-core/internal/abort';

const ABORT_MESSAGE = 'DOCX document session was aborted';

const getDocxWasmModule = createLazyWasmModule(() => resolveWasm(
    import.meta.url,
    'docx_parser_bg.wasm',
    '@silurus/ooxml-docx/wasm-binary',
  ));

/** Options for the bounded Node DOCX page session. */
export interface OpenDocxDocumentOptions extends OoxmlNodeSessionOptions {
  /**
   * Opt-in model source for document input that is not an OOXML package (a
   * `ModelSource<'docx'>`). When its `claim()` accepts the bytes, it opens them
   * in this realm; otherwise the ordinary OOXML path runs unchanged.
   */
  legacyDoc?: ModelSource<'docx'>;
  /**
   * Opt-in model source for OpenDocument text input (a `ModelSource<'docx'>`, such as the
   * one an OpenDocument companion package exports). It is asked after
   * `legacyDoc`; when its `claim()` accepts the raw bytes it opens them through its
   * own module, otherwise the ordinary OOXML path runs unchanged. A source for
   * another format is a `TypeError`. Omitting this option imports and runs
   * nothing extra.
   */
  odt?: ModelSource<'docx'>;
  cjkFallback?: CjkFallback;
  /** Canvas implementation used for text measurement and page allocation. */
  factory: NodeCanvasFactory;
  /** Stable DATE/TIME field instant captured before pagination. */
  currentDate?: Date | number;
  /** Continue paragraph footnotes across physical pages. Omitted means true,
   * the same library display default as the browser `LoadOptions`: a policy
   * chosen to follow Word's page allocation in finite controls, not a
   * normative ECMA-376 rule. Pass `false` to keep the previous whole-note
   * pagination. With continuation, table splitting, width changes and changed
   * source cuts can reject layout, and full-note acquisition is subject to the
   * library source/page budgets. */
  allowFootnoteContinuation?: boolean;
}

export interface DocxPageRenderOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  onTextRun?: (run: DocxTextRunInfo) => void;
}

export interface DocxRenderedPage {
  readonly pageIndex: number;
  readonly widthPt: number;
  readonly heightPt: number;
  /** Caller-owned after yield; retaining it is outside the session memory bound. */
  readonly canvas: NodeCanvasLike;
}

/**
 * Read-only Node DOCX session for the pictures a document model names. It
 * needs no canvas factory: nothing is measured, paginated or rendered, so a
 * process that only extracts content need not load a canvas backend.
 */
export interface DocxPictureSession {
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  /**
   * Raw bytes of one picture the document model names (an image run's
   * `imagePath`), as a Blob of `mimeType` (default
   * `application/octet-stream`). For a `.docx` any picture part is served (a
   * part whose content type is `image/*`, such as `word/media/image1.png`);
   * for an application model source, the picture keys its model assigns.
   * Another path rejects with a `TypeError` (or the archive's not-found
   * error) and leaves the session usable. Reads charge the session's archive
   * budgets, results are kept in a bounded raw-part cache (shared with page
   * rendering in a {@link DocxDocumentSession}), a resource-limit failure
   * fails the session, and the call rejects after `close()`.
   */
  getImage(path: string, mimeType?: string): Promise<Blob>;
  close(): Promise<void>;
}

/** Options for {@link openDocxPictures}: those of `openDocxDocument` that do
 * not concern measurement, pagination or rendering. */
export type OpenDocxPicturesOptions =
  OoxmlNodeSessionOptions & Pick<OpenDocxDocumentOptions, 'legacyDoc' | 'odt'>;

/**
 * Ready, paginated Node document. DOCX pagination is necessarily sequential and
 * retained because preceding flow determines later pages and total page count.
 * Parsing and transfer are nevertheless bounded: no whole document XML/JSON
 * value crosses the cursor, and only one page canvas is created per pull.
 */
export interface DocxDocumentSession
  extends DocxPictureSession, AsyncIterable<DocxRenderedPage> {
  readonly pageCount: number;
  pageSize(pageIndex: number): Readonly<{ widthPt: number; heightPt: number }>;
  renderPage(pageIndex: number, options?: DocxPageRenderOptions): Promise<NodeCanvasLike>;
  pages(options?: DocxPageRenderOptions): AsyncGenerator<DocxRenderedPage, void, void>;
}

/**
 * Open a Node DOCX session that parses through the same acknowledged body cursor
 * as the browser Viewer, seals the canonical layout source, completes pagination,
 * and then renders one caller-owned canvas at a time.
 */
export async function openDocxDocument(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenDocxDocumentOptions,
): Promise<DocxDocumentSession> {
  if (!options?.factory) throw new TypeError(
    'openDocxDocument requires a canvas factory; openDocxPictures reads pictures without one',
  );
  if (__OOXML_MODEL_SOURCES__ && (options.legacyDoc !== undefined || options.odt !== undefined)) {
    const { openDocxSource } = await import('./docx-model-source.ts');
    return openDocxSource(buffer, options, getDocxWasmModule);
  }
  const cjkFallback = resolveCjkFallback(options.cjkFallback);
  const acquired = await acquireDocxNodeDocument(
    toUint8(buffer),
    getDocxWasmModule(),
    options,
    (transport, identity, pullOptions) =>
      materializeDocumentPullLayoutSession(transport, identity, pullOptions),
  );
  try {
    throwIfAborted(options.signal, ABORT_MESSAGE);
    const measurementCanvas = options.factory.createCanvas(1, 1);
    const services = createLayoutServices(acquired.result, {
      cjkFallback,
      measureContext: measurementCanvas.getContext('2d') as CanvasRenderingContext2D,
      // Symbol colour probes draw on this session's own canvas backend.
      createCanvas: (width: number, height: number) => options.factory.createCanvas(width, height) as never,
      // Library default: continuation unless the caller passes exactly false.
      allowFootnoteContinuation: options.allowFootnoteContinuation !== false,
    });
    const defaultCurrentDateMs = normalizeCurrentDate(options.currentDate);
    const retained = retainRenderWorkerDocumentLayout(
      acquired.result,
      services,
      defaultCurrentDateMs,
    );
    const layout = retained.layoutVariants.defaultLayout;
    const session = new DocxDocumentSessionImpl(
      acquired.closeArchive,
      acquired.archive as unknown as DocxNodeArchive,
      acquired.result,
      services,
      layout,
      options.factory,
      defaultCurrentDateMs,
      acquired.usage,
      acquired.metrics,
      options.signal,
    );
    acquired.metrics.observeUsage(session.resourceUsage);
    acquired.metrics.checkpoint('pagination ready');
    return session;
  } catch (error) {
    try {
      acquired.closeArchive();
    } catch {
      // Preserve the layout failure.
    }
    const normalized = parseResourceLimitError(error) ?? error;
    acquired.metrics.fail(normalized);
    throw normalized;
  }
}

/** Materialize the public DOCX compatibility model through the acknowledged
 * body-unit coordinator without creating measurement or page-layout state. */
export async function materializeDocxDocument(
  buffer: ArrayBuffer | Uint8Array,
  options: OoxmlNodeSessionOptions & Pick<OpenDocxDocumentOptions, 'legacyDoc' | 'odt'> = {},
): Promise<DocxDocumentModel> {
  if (__OOXML_MODEL_SOURCES__ && (options.legacyDoc !== undefined || options.odt !== undefined)) {
    const { materializeDocxSource } = await import('./docx-model-source.ts');
    return materializeDocxSource(buffer, options, getDocxWasmModule);
  }
  return usingOwnedSession(
    async () => {
      const acquired = await acquireDocxNodeDocument(
        toUint8(buffer),
        getDocxWasmModule(),
        options,
        (transport, identity, pullOptions) =>
          materializeDocumentPullSession(transport, identity, pullOptions),
      );
      let succeeded = false;
      return {
        acquired,
        markSucceeded: () => { succeeded = true; },
        close: async () => {
          try {
            acquired.closeArchive();
            if (succeeded) acquired.metrics.succeed({ documents: 1 });
          } catch (error) {
            acquired.metrics.fail(error);
            throw error;
          }
        },
      };
    },
    async ({ acquired, markSucceeded }) => {
      try {
        const document = normalizeDocxDocumentModel(acquired.result);
        acquired.metrics.checkpoint('document materialized', acquired.usage);
        markSucceeded();
        return document;
      } catch (error) {
        acquired.metrics.fail(error);
        throw error;
      }
    },
  );
}

/**
 * Open a read-only session over the pictures of a DOCX document (or of a
 * `legacyDoc` / `odt` source) without a canvas factory. The document is parsed
 * through the same acknowledged body cursor and archive budgets as
 * {@link openDocxDocument}, which validates it and lets a model source build
 * its picture store, but no layout state is created. Pair it with
 * {@link materializeDocxDocument} for the `imagePath`s to read.
 */
export async function openDocxPictures(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenDocxPicturesOptions = {},
): Promise<DocxPictureSession> {
  if (__OOXML_MODEL_SOURCES__ && (options.legacyDoc !== undefined || options.odt !== undefined)) {
    const { openDocxPicturesSource } = await import('./docx-model-source.ts');
    return openDocxPicturesSource(buffer, options, getDocxWasmModule);
  }
  const acquired = await acquireDocxNodeDocument(
    toUint8(buffer),
    getDocxWasmModule(),
    options,
    (transport, identity, pullOptions) =>
      materializeDocumentPullSession(transport, identity, pullOptions),
  );
  return admitDocxPictureSession(acquired, options.signal);
}

/** Wrap an acquired (already parsed) document in a picture session that owns
 * its archive from here on, including on failure. */
export function admitDocxPictureSession(
  acquired: Readonly<{
    archive: unknown;
    usage: OoxmlResourceUsageSnapshot | undefined;
    metrics: OoxmlResourceMetricsSession;
    closeArchive(): void;
  }>,
  signal: AbortSignal | undefined,
): DocxPictureSession {
  try {
    throwIfAborted(signal, ABORT_MESSAGE);
    const session = new DocxPictureSessionImpl(
      acquired.closeArchive,
      acquired.archive as DocxNodeArchive,
      acquired.usage,
      acquired.metrics,
      signal,
    );
    acquired.metrics.observeUsage(session.resourceUsage);
    acquired.metrics.checkpoint('pictures ready');
    return session;
  } catch (error) {
    try {
      acquired.closeArchive();
    } catch {
      // Preserve the admission failure.
    }
    const normalized = parseResourceLimitError(error) ?? error;
    acquired.metrics.fail(normalized);
    throw normalized;
  }
}

type SessionState = Readonly<{
  source: Awaited<ReturnType<typeof materializeDocumentPullLayoutSession>>;
  services: ReturnType<typeof createLayoutServices>;
}>;

type DefaultDocumentLayout =
  ReturnType<typeof retainRenderWorkerDocumentLayout>['layoutVariants']['defaultLayout'];

/**
 * Picture reads, resource accounting and archive ownership. Used alone for a
 * factory-less {@link DocxPictureSession} and as the base of the rendering
 * {@link DocxDocumentSessionImpl}, so both serve `getImage()` identically.
 */
export class DocxPictureSessionImpl implements DocxPictureSession {
  protected lastResourceUsage: OoxmlResourceUsageSnapshot | undefined;
  protected closed = false;
  private closePromise: Promise<void> | undefined;
  /** Set just before the archive is freed: an accepted render may still read,
   * but nothing reaches a freed archive. */
  private released = false;
  private resourceFailure: OoxmlResourceLimitError | null = null;
  /** Page rendering and getImage() share one retained byte owner per part,
   * bounded like the PPTX session's raw-part cache. */
  private readonly rawParts = new BoundedRawPartCache({
    maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
    maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
  });
  protected readonly fetchImage = (path: string, mimeType: string): Promise<Blob> =>
    this.rawParts.get(path, mimeType, () => {
      if (this.released) throw new Error('DOCX document session is closed');
      throwIfAborted(this.signal, ABORT_MESSAGE);
      const bytes = this.archive.extract_image(path);
      this.refreshResourceUsage();
      return new Blob([bytes as BlobPart], { type: mimeType });
    });

  constructor(
    private readonly closeArchive: () => void,
    private readonly archive: DocxNodeArchive,
    usage: OoxmlResourceUsageSnapshot | undefined,
    protected readonly metrics: OoxmlResourceMetricsSession,
    protected readonly signal?: AbortSignal,
  ) {
    this.lastResourceUsage = usage;
  }

  get resourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    if (this.closed) return this.lastResourceUsage;
    return this.refreshResourceUsage();
  }

  private refreshResourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    try {
      this.lastResourceUsage = decodeOoxmlResourceUsage(this.archive.resource_usage());
      this.metrics.observeUsage(this.lastResourceUsage);
    } catch {
      // A closed/trapped archive cannot improve the last valid diagnostic
      // checkpoint. Rendering failures remain surfaced by their operation.
    }
    return this.lastResourceUsage;
  }

  async getImage(path: string, mimeType = 'application/octet-stream'): Promise<Blob> {
    this.assertOpen();
    try {
      const picture = this.archive.is_picture_part?.(path);
      // The first check reads [Content_Types].xml: account for it even when
      // the path is then refused.
      if (picture !== undefined) this.refreshResourceUsage();
      if (picture === false) throw new TypeError(NOT_A_PICTURE);
      return await this.fetchImage(path, mimeType);
    } catch (error) {
      // Only a resource-limit overrun fails the session; a non-picture or
      // missing path rejects this call alone.
      const normalized = parseResourceLimitError(error) ?? error;
      if (normalized instanceof OoxmlResourceLimitError) this.failOperation(normalized);
      throw normalized;
    }
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.refreshResourceUsage();
    this.closed = true;
    this.closePromise = this.release();
    return this.closePromise;
  }

  /** Terminal outcome reported to the resource metrics on a clean close. */
  protected outcome(): Readonly<Record<string, number>> {
    return { documents: 1 };
  }

  protected async release(): Promise<void> {
    dropDecodedBitmapCache(this.fetchImage);
    dropSvgImageCache(this.fetchImage);
    this.rawParts.clear();
    this.released = true;
    try {
      this.closeArchive();
    } catch (error) {
      const normalized = parseResourceLimitError(error) ?? error;
      this.metrics.fail(normalized);
      throw normalized;
    }
    this.metrics.checkpoint('document session closed', this.lastResourceUsage);
    this.metrics.succeed(this.outcome());
  }

  protected assertOpen(): void {
    if (this.closed) throw new Error('DOCX document session is closed');
    if (this.resourceFailure) throw this.resourceFailure;
  }

  protected failOperation(error: unknown): never {
    const normalized = parseResourceLimitError(error) ?? error;
    if (normalized instanceof OoxmlResourceLimitError) {
      this.resourceFailure ??= normalized;
    }
    this.metrics.fail(normalized);
    throw normalized;
  }
}

export class DocxDocumentSessionImpl extends DocxPictureSessionImpl implements DocxDocumentSession {
  readonly pageCount: number;
  private readonly sizes: ReadonlyArray<Readonly<{ widthPt: number; heightPt: number }>>;
  private state: SessionState | null;
  private renderTail: Promise<void> = Promise.resolve();
  private pagesStarted = false;

  constructor(
    closeArchive: () => void,
    archive: DocxNodeArchive,
    source: SessionState['source'],
    services: SessionState['services'],
    layout: DefaultDocumentLayout,
    private readonly factory: NodeCanvasFactory,
    private readonly defaultCurrentDateMs: number,
    usage: OoxmlResourceUsageSnapshot | undefined,
    metrics: OoxmlResourceMetricsSession,
    signal?: AbortSignal,
  ) {
    super(closeArchive, archive, usage, metrics, signal);
    this.state = { source, services };
    this.pageCount = layout.pages.length;
    this.sizes = Object.freeze(layout.pages.map((page) => Object.freeze({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })));
  }

  pageSize(pageIndex: number): Readonly<{ widthPt: number; heightPt: number }> {
    const size = this.sizes[pageIndex];
    if (!size) throw new RangeError(`DOCX page index ${pageIndex} out of range`);
    return size;
  }

  [Symbol.asyncIterator](): AsyncGenerator<DocxRenderedPage, void, void> {
    return this.pages();
  }

  renderPage(pageIndex: number, options: DocxPageRenderOptions = {}): Promise<NodeCanvasLike> {
    try {
      this.assertOpen();
    } catch (error) {
      return Promise.reject(error);
    }
    this.pageSize(pageIndex);
    return this.enqueueRender(async () => {
      throwIfAborted(this.signal, ABORT_MESSAGE);
      const state = this.requireState();
      const canvas = this.factory.createCanvas(1, 1);
      await withNodeCanvasRuntime(this.factory, () => renderLayoutSourceToCanvas(
        state.source,
        canvas as unknown as HTMLCanvasElement,
        pageIndex,
        {
          ...options,
          currentDate: this.defaultCurrentDateMs,
          defaultCurrentDateMs: this.defaultCurrentDateMs,
          layoutServices: state.services,
          fetchImage: this.fetchImage,
        },
      ));
      throwIfAborted(this.signal, ABORT_MESSAGE);
      return canvas;
    }).catch((error: unknown) => this.failOperation(error));
  }

  async *pages(options: DocxPageRenderOptions = {}): AsyncGenerator<DocxRenderedPage, void, void> {
    if (this.closed) throw new Error('DOCX document session is closed');
    if (this.pagesStarted) throw new Error('DOCX page stream is one-pass and was already consumed');
    this.pagesStarted = true;
    let operationError: unknown;
    try {
      for (let pageIndex = 0; pageIndex < this.pageCount; pageIndex += 1) {
        const canvas = await this.renderPage(pageIndex, options);
        const size = this.pageSize(pageIndex);
        yield { pageIndex, ...size, canvas };
      }
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
      throw operationError;
    } finally {
      try {
        await this.close();
      } catch (cleanupError) {
        if (operationError === undefined) throw cleanupError;
      }
    }
  }

  protected override outcome(): Readonly<Record<string, number>> {
    return { pages: this.pageCount };
  }

  protected override async release(): Promise<void> {
    await this.renderTail;
    this.state = null;
    return super.release();
  }

  private enqueueRender<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.renderTail.then(operation, operation);
    this.renderTail = result.then(() => undefined, () => undefined);
    return result;
  }

  private requireState(): SessionState {
    if (!this.state) throw new Error('DOCX document session is closed');
    return this.state;
  }
}

const NOT_A_PICTURE = 'DOCX getImage(): the path does not name a picture part (a part whose content type is image/*)';

declare const __OOXML_MODEL_SOURCES__: boolean;
