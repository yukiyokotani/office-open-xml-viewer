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
  type OoxmlResourceMetricsSession,
} from '@silurus/ooxml-core/worker';
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
import {
  type NodeCanvasFactory,
  type NodeCanvasLike,
  withNodeCanvasRuntime,
} from './render.ts';
import { createLazyWasmModule, resolveWasm } from './wasm-loader.ts';
import { usingOwnedSession } from '@silurus/ooxml-core/internal/owned-session';
import { normalizeNodeOfficeInput } from './normalize-input.ts';

const getDocxWasmModule = createLazyWasmModule(() => resolveWasm(
    import.meta.url,
    'docx_parser_bg.wasm',
    '@silurus/ooxml-docx/wasm-binary',
  ));

/** Options for the bounded Node DOCX page session. */
export interface OpenDocxDocumentOptions extends OoxmlNodeSessionOptions {
  /** Canvas implementation used for text measurement and page allocation. */
  factory: NodeCanvasFactory;
  /** Stable DATE/TIME field instant captured before pagination. */
  currentDate?: Date | number;
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
 * Ready, paginated Node document. DOCX pagination is necessarily sequential and
 * retained because preceding flow determines later pages and total page count.
 * Parsing and transfer are nevertheless bounded: no whole document XML/JSON
 * value crosses the cursor, and only one page canvas is created per pull.
 */
export interface DocxDocumentSession extends AsyncIterable<DocxRenderedPage> {
  readonly pageCount: number;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  pageSize(pageIndex: number): Readonly<{ widthPt: number; heightPt: number }>;
  renderPage(pageIndex: number, options?: DocxPageRenderOptions): Promise<NodeCanvasLike>;
  pages(options?: DocxPageRenderOptions): AsyncGenerator<DocxRenderedPage, void, void>;
  close(): Promise<void>;
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
  if (!options?.factory) throw new TypeError('openDocxDocument requires a canvas factory');
  const normalized = await normalizeNodeOfficeInput(buffer, 'docx', options);
  const acquired = await acquireDocxNodeDocument(
    normalized,
    getDocxWasmModule(),
    options,
    (transport, identity, pullOptions) =>
      materializeDocumentPullLayoutSession(transport, identity, pullOptions),
  );
  try {
    throwIfAborted(options.signal);
    const measurementCanvas = options.factory.createCanvas(1, 1);
    const services = createLayoutServices(acquired.result, {
      measureContext: measurementCanvas.getContext('2d') as CanvasRenderingContext2D,
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
      acquired.archive,
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
  options: OoxmlNodeSessionOptions = {},
): Promise<DocxDocumentModel> {
  const normalized = await normalizeNodeOfficeInput(buffer, 'docx', options);
  return usingOwnedSession(
    async () => {
      const acquired = await acquireDocxNodeDocument(
        normalized,
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

type SessionState = Readonly<{
  source: Awaited<ReturnType<typeof materializeDocumentPullLayoutSession>>;
  services: ReturnType<typeof createLayoutServices>;
}>;

type DefaultDocumentLayout =
  ReturnType<typeof retainRenderWorkerDocumentLayout>['layoutVariants']['defaultLayout'];

class DocxDocumentSessionImpl implements DocxDocumentSession {
  readonly pageCount: number;
  private readonly sizes: ReadonlyArray<Readonly<{ widthPt: number; heightPt: number }>>;
  private lastResourceUsage: OoxmlResourceUsageSnapshot | undefined;
  private state: SessionState | null;
  private renderTail: Promise<void> = Promise.resolve();
  private pagesStarted = false;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private resourceFailure: OoxmlResourceLimitError | null = null;
  private readonly fetchImage = async (path: string, mimeType: string): Promise<Blob> => {
    const bytes = this.archive.extract_image(path);
    return new Blob([bytes as BlobPart], { type: mimeType });
  };

  constructor(
    private readonly closeArchive: () => void,
    private readonly archive: DocxNodeArchive,
    source: SessionState['source'],
    services: SessionState['services'],
    layout: DefaultDocumentLayout,
    private readonly factory: NodeCanvasFactory,
    private readonly defaultCurrentDateMs: number,
    usage: OoxmlResourceUsageSnapshot | undefined,
    private readonly metrics: OoxmlResourceMetricsSession,
    private readonly signal?: AbortSignal,
  ) {
    this.state = { source, services };
    this.pageCount = layout.pages.length;
    this.lastResourceUsage = usage;
    this.sizes = Object.freeze(layout.pages.map((page) => Object.freeze({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })));
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

  pageSize(pageIndex: number): Readonly<{ widthPt: number; heightPt: number }> {
    const size = this.sizes[pageIndex];
    if (!size) throw new RangeError(`DOCX page index ${pageIndex} out of range`);
    return size;
  }

  [Symbol.asyncIterator](): AsyncGenerator<DocxRenderedPage, void, void> {
    return this.pages();
  }

  renderPage(pageIndex: number, options: DocxPageRenderOptions = {}): Promise<NodeCanvasLike> {
    if (this.closed) return Promise.reject(new Error('DOCX document session is closed'));
    if (this.resourceFailure) return Promise.reject(this.resourceFailure);
    this.pageSize(pageIndex);
    return this.enqueueRender(async () => {
      throwIfAborted(this.signal);
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
      throwIfAborted(this.signal);
      return canvas;
    }).catch((error: unknown) => {
      const normalized = parseResourceLimitError(error) ?? error;
      if (normalized instanceof OoxmlResourceLimitError) {
        this.resourceFailure ??= normalized;
      }
      this.metrics.fail(normalized);
      throw normalized;
    });
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

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.refreshResourceUsage();
    this.closed = true;
    this.closePromise = this.release();
    return this.closePromise;
  }

  private enqueueRender<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.renderTail.then(operation, operation);
    this.renderTail = result.then(() => undefined, () => undefined);
    return result;
  }

  private async release(): Promise<void> {
    await this.renderTail;
    dropDecodedBitmapCache(this.fetchImage);
    dropSvgImageCache(this.fetchImage);
    this.state = null;
    try {
      this.closeArchive();
    } catch (error) {
      const normalized = parseResourceLimitError(error) ?? error;
      this.metrics.fail(normalized);
      throw normalized;
    }
    this.metrics.checkpoint('document session closed', this.lastResourceUsage);
    this.metrics.succeed({ pages: this.pageCount });
  }

  private requireState(): SessionState {
    if (!this.state) throw new Error('DOCX document session is closed');
    return this.state;
  }
}

function normalizeCurrentDate(value: Date | number | undefined): number {
  const current = value instanceof Date ? value.getTime() : (value ?? Date.now());
  if (!Number.isFinite(current)) {
    throw new RangeError('currentDate must resolve to finite epoch milliseconds');
  }
  return current;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('DOCX document session was aborted');
  error.name = 'AbortError';
  throw error;
}
