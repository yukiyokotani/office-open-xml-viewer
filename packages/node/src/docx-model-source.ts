import {
  acquireDocxNodeDocument,
  createLayoutServices,
  materializeDocumentPullLayoutSession,
  materializeDocumentPullSession,
  normalizeDocxDocumentModel,
  normalizeDocxParserError,
  normalizeLayoutOptions,
  retainRenderWorkerDocumentLayout,
  type AcquiredDocxNodeDocument,
  type DocxNodeAcquisitionOptions,
  type DocxNodePullIdentity,
  type DocxNodePullOptions,
  type DocxNodePullTransport,
  type DocxNodeArchive,
} from '@silurus/ooxml-docx/internal/session';
import { resolveCjkFallback } from '@silurus/ooxml-core';
import { parseResourceLimitError, parseTypedParserError } from '@silurus/ooxml-core/worker';
import { usingOwnedSession } from '@silurus/ooxml-core/internal/owned-session';
import type { DocxDocumentModel } from '@silurus/ooxml-docx';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import { resolveNodeSessionInput } from './model-source.ts';
import {
  DocxDocumentSessionImpl,
  type DocxDocumentSession,
  type DocxPageRenderOptions,
  type OpenDocxDocumentOptions,
} from './docx.ts';

class SourceDocxDocumentSession extends DocxDocumentSessionImpl {
  constructor(
    ...args: ConstructorParameters<typeof DocxDocumentSessionImpl>
  ) {
    super(...args);
  }

  override renderPage(pageIndex: number, options: DocxPageRenderOptions = {}) {
    return super.renderPage(pageIndex, { ...options, showTrackedChanges: true } as DocxPageRenderOptions);
  }
}

export async function openDocxSource(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenDocxDocumentOptions,
  wasmModule: () => WebAssembly.Module,
): Promise<DocxDocumentSession> {
  const cjkFallback = resolveCjkFallback(options.cjkFallback);
  const { acquired, viewDefaults } = await acquireDocxInput(
    buffer, options, wasmModule, materializeDocumentPullLayoutSession,
  );
  try {
    if (options.signal?.aborted) {
      const error = new Error('DOCX document session was aborted');
      error.name = 'AbortError';
      throw error;
    }
    const measurementCanvas = options.factory.createCanvas(1, 1);
    const services = createLayoutServices(acquired.result, {
      cjkFallback,
      measureContext: measurementCanvas.getContext('2d') as CanvasRenderingContext2D,
      // Same library default as openDocxDocument's OOXML branch.
      allowFootnoteContinuation: options.allowFootnoteContinuation !== false,
    });
    const defaultCurrentDateMs = normalizeCurrentDate(options.currentDate);
    const retained = retainRenderWorkerDocumentLayout(acquired.result, services, defaultCurrentDateMs);
    const showTrackedChanges = viewDefaults.showTrackedChanges === true;
    const layout = showTrackedChanges
      ? retained.layoutVariants.layoutFor(
        normalizeLayoutOptions(defaultCurrentDateMs, defaultCurrentDateMs, true),
      )
      : retained.layoutVariants.defaultLayout;
    const Session = showTrackedChanges ? SourceDocxDocumentSession : DocxDocumentSessionImpl;
    const session = new Session(
      acquired.closeArchive, acquired.archive as unknown as DocxNodeArchive, acquired.result, services,
      layout, options.factory, defaultCurrentDateMs, acquired.usage,
      acquired.metrics, options.signal,
    );
    acquired.metrics.observeUsage(session.resourceUsage);
    acquired.metrics.checkpoint('pagination ready');
    return session;
  } catch (error) {
    try { acquired.closeArchive(); } catch {}
    const normalized = parseResourceLimitError(error) ?? error;
    acquired.metrics.fail(normalized);
    throw normalized;
  }
}

export async function materializeDocxSource(
  buffer: ArrayBuffer | Uint8Array,
  options: OoxmlNodeSessionOptions & Pick<OpenDocxDocumentOptions, 'legacyDoc' | 'odt'>,
  wasmModule: () => WebAssembly.Module,
): Promise<DocxDocumentModel> {
  return usingOwnedSession(
    async () => {
      const { acquired } = await acquireDocxInput(
        buffer, options, wasmModule, materializeDocumentPullSession,
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

function normalizeCurrentDate(value: Date | number | undefined): number {
  const current = value instanceof Date ? value.getTime() : (value ?? Date.now());
  if (!Number.isFinite(current)) throw new RangeError('currentDate must resolve to finite epoch milliseconds');
  return current;
}

/** Selected-source DOCX acquisition, loaded only after source-option dispatch. */
export async function acquireDocxInput<TResult>(
  buffer: ArrayBuffer | Uint8Array,
  options: OoxmlNodeSessionOptions & DocxNodeAcquisitionOptions & Pick<OpenDocxDocumentOptions, 'legacyDoc' | 'odt'>,
  wasmModule: () => WebAssembly.Module,
  consume: (
    transport: DocxNodePullTransport,
    identity: DocxNodePullIdentity,
    options: DocxNodePullOptions,
  ) => Promise<TResult>,
): Promise<Readonly<{
  acquired: AcquiredDocxNodeDocument<TResult>;
  viewDefaults: Readonly<{ showTrackedChanges?: boolean }>;
}>> {
  const {
    acquireDocxSessionFromArchive,
    validateDocxModelSourceArchive,
    validateDocxModelSourceViewDefaults,
  } = await import('@silurus/ooxml-docx/internal/model-source-session');
  const input = await resolveNodeSessionInput(
    buffer, 'docx', { ...options, sources: [options.legacyDoc, options.odt] }, validateDocxModelSourceArchive,
    error => parseTypedParserError(error) ?? normalizeDocxParserError(error),
  );
  if (input.kind === 'ooxml') {
    return {
      acquired: await acquireDocxNodeDocument(input.bytes, wasmModule(), options, consume),
      viewDefaults: {},
    };
  }
  let viewDefaults: Readonly<{ showTrackedChanges?: boolean }>;
  try {
    viewDefaults = validateDocxModelSourceViewDefaults(input.opened.viewDefaults);
  } catch (error) {
    try { input.opened.close(); } catch {}
    throw error;
  }
  const acquired = await acquireDocxSessionFromArchive({
    archive: input.opened.archive,
    sourceByteLength: input.sourceByteLength,
    closeArchive: input.opened.close,
  }, options, consume);
  return { acquired, viewDefaults };
}
