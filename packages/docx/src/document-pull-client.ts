import type { OoxmlResourceUsageSnapshot, WorkerBridgeTransport } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
  PULL_SESSION_PROTOCOL,
  pullWithCreditRetry,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { BodyElement, DocxDocumentModel } from './types.js';
import {
  layoutSourceModelAdapterFromOwnedModel,
  layoutSourceStoreFromOwnedModel,
  type LayoutSourceModelAdapter,
} from './layout-source-model-adapter.js';
import { errorPageLayoutReady } from './layout/document-layout-variants.js';
import type { LayoutSourceStore } from './layout/layout-source-store.js';

const DOCX_INITIAL_BODY_PULL_BYTES = 1024 * 1024;
const MAX_DOCUMENT_UNIT_BYTES = Math.max(
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
);

type DocumentUnit =
  | { readonly kind: 'body'; readonly body: BodyElement[] }
  | { readonly kind: 'complete'; readonly document: DocxDocumentModel };

export interface MaterializeDocumentPullOptions {
  readonly timeoutMs?: number;
  readonly signal?: AbortSignal;
  /**
   * The owner's lifetime, for the wait that follows the pull of a document
   * that failed to parse (its error page's code, #2248): ending the transport
   * settles a pull, but not that wait. Read only for such a document.
   * Defaults to `signal`.
   */
  readonly lifetime?: () => AbortSignal | undefined;
  readonly onUsage?: (usage: OoxmlResourceUsageSnapshot) => void;
}

interface DocumentPullAccumulator<TResult> {
  acceptBody(body: BodyElement[]): void;
  complete(document: DocxDocumentModel): TResult;
  /** Code the result needs before it is handed out; `undefined` when there is
   * nothing to wait for (always, for a document that parsed). */
  ready?(result: TResult): Promise<unknown> | undefined;
}

/** Drain one sequential DOCX operation into the backward-compatible public
 * model. Each transferred JSON unit is decoded, accepted, and ACKed before the
 * worker may produce another; no monolithic document JSON crosses realms. */
export async function materializeDocumentPullSession(
  transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>,
  identity: PullSessionIdentity<number>,
  options: MaterializeDocumentPullOptions = {},
): Promise<DocxDocumentModel> {
  const body: BodyElement[] = [];
  return drainDocumentPullSession(transport, identity, options, {
    acceptBody: (unitBody) => { body.push(...unitBody); },
    complete: (document) => {
      document.body = body;
      return document;
    },
  });
}

/** Drain directly into the immutable layout source used by Node rendering.
 * Transferred units already have exclusive ownership, so the layout sink can
 * consume them without constructing the public compatibility graph or cloning
 * every body unit. */
export async function materializeDocumentPullLayoutSession(
  transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>,
  identity: PullSessionIdentity<number>,
  options: MaterializeDocumentPullOptions = {},
): Promise<LayoutSourceStore> {
  const ownedBody: BodyElement[] = [];
  return drainDocumentPullSession(transport, identity, options, {
    acceptBody: (body) => { ownedBody.push(...body); },
    complete: (document) => {
      document.body = ownedBody;
      return layoutSourceStoreFromOwnedModel(document);
    },
    // A fatally unparseable document loads its error page here (#2248).
    ready: (source: LayoutSourceStore) => errorPageLayoutReady(source, options.lifetime?.() ?? options.signal),
  });
}

/** Drain a sequential DOCX operation directly into its two required ownership
 * graphs: the mutable public compatibility model and the immutable layout
 * source. The second graph is detached while each bounded body unit is live,
 * avoiding a full-document clone after the complete public body accumulates. */
export async function materializeDocumentPullAdapterSession(
  transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>,
  identity: PullSessionIdentity<number>,
  options: MaterializeDocumentPullOptions = {},
): Promise<LayoutSourceModelAdapter> {
  const models = await materializeDocumentPullOwnedModelsSession(transport, identity, options);
  const adapter = layoutSourceModelAdapterFromOwnedModel(models.document, models.ownedLayoutDocument);
  // A fatally unparseable document loads its error page here (#2248).
  const errorPage = errorPageLayoutReady(adapter.source, options.lifetime?.() ?? options.signal);
  if (errorPage) await errorPage;
  return adapter;
}

export interface MaterializedDocumentPullOwnedModels {
  /** Parser-shaped model retained for public compatibility or worker fallback. */
  readonly document: DocxDocumentModel;
  /** Disjoint, builder-owned graph that may be consumed only by the adapter. */
  readonly ownedLayoutDocument: DocxDocumentModel;
}

/** Lower-level acquisition boundary for render workers that must inspect the
 * parser-shaped compatibility graph before deciding whether Window fallback is
 * required. Callers must either consume `ownedLayoutDocument` with
 * `layoutSourceModelAdapterFromOwnedModel` or release it. */
export async function materializeDocumentPullOwnedModelsSession(
  transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>,
  identity: PullSessionIdentity<number>,
  options: MaterializeDocumentPullOptions = {},
): Promise<MaterializedDocumentPullOwnedModels> {
  const publicBody: BodyElement[] = [];
  const ownedLayoutBody: BodyElement[] = [];
  return drainDocumentPullSession(transport, identity, options, {
    acceptBody: (body) => {
      // Clone only the current credit-bounded unit. The parsed instances become
      // the public model; their detached counterparts become the builder-owned
      // layout graph and are never exposed to callers.
      const layoutUnit = structuredClone(body);
      for (const element of body) publicBody.push(element);
      for (const element of layoutUnit) ownedLayoutBody.push(element);
    },
    complete: (document) => {
      // The terminal envelope has its own hard byte cap and carries no body.
      // Detach it before attaching either accumulated graph.
      const ownedLayoutDocument = structuredClone(document);
      document.body = publicBody;
      ownedLayoutDocument.body = ownedLayoutBody;
      return Object.freeze({ document, ownedLayoutDocument });
    },
  });
}

/** The sole DOCX pull-protocol state machine. Ownership policy is injected by
 * the accumulator so ACK, validation, usage, cancellation, and transfer cleanup
 * cannot drift between compatibility and layout-aware materialization paths. */
async function drainDocumentPullSession<TResult>(
  transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>,
  identity: PullSessionIdentity<number>,
  options: MaterializeDocumentPullOptions,
  accumulator: DocumentPullAccumulator<TResult>,
): Promise<TResult> {
  const session = new BoundedPullSession(transport, {
    ...identity,
    maxByteCredit: MAX_DOCUMENT_UNIT_BYTES,
    timeoutMs: options.timeoutMs,
  });
  try {
    for (;;) {
      const chunk = await pullWithCreditRetry(
        session,
        DOCX_INITIAL_BODY_PULL_BYTES,
        MAX_DOCUMENT_UNIT_BYTES,
        options.signal,
      );
      try {
        const usage = chunk.usage ?? session.usageCheckpoint;
        if (usage) options.onUsage?.(usage);
        const unit = parseDocumentUnit(chunk.payload);
        if (chunk.done !== (unit.kind === 'complete')) {
          throw new TypeError('DOCX document unit terminal flag does not match its payload');
        }
        if (unit.kind === 'body') {
          accumulator.acceptBody(unit.body);
          await chunk.ack({ signal: options.signal });
          continue;
        }
        if (!Array.isArray(unit.document.body) || unit.document.body.length !== 0) {
          throw new TypeError('DOCX terminal document must not duplicate streamed body blocks');
        }
        const result = accumulator.complete(unit.document);
        const ready = accumulator.ready?.(result);
        if (ready) await ready;
        await chunk.ack({ signal: options.signal });
        return result;
      } finally {
        chunk.disposeTransferred();
      }
    }
  } catch (error) {
    await session.cancel('request-error').catch(() => undefined);
    throw error;
  }
}

export function isDocumentPullResponse(
  value: unknown,
): value is PullSessionResponse<ArrayBuffer, number> {
  return !!value && typeof value === 'object'
    && (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

function parseDocumentUnit(payload: ArrayBuffer): DocumentUnit {
  const value = JSON.parse(new TextDecoder().decode(new Uint8Array(payload))) as unknown;
  if (!value || typeof value !== 'object') throw new TypeError('DOCX document unit must be an object');
  const record = value as Record<string, unknown>;
  if (record.kind === 'body' && Array.isArray(record.body)) {
    return record as unknown as Extract<DocumentUnit, { kind: 'body' }>;
  }
  if (record.kind === 'complete' && record.document && typeof record.document === 'object') {
    return record as unknown as Extract<DocumentUnit, { kind: 'complete' }>;
  }
  throw new TypeError('DOCX document unit has an unknown shape');
}
