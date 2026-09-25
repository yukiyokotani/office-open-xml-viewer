import {
  decodeOoxmlResourceUsage,
  exactTransferableArrayBuffer,
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
  PULL_SESSION_PROTOCOL,
  PullSessionInsufficientCreditError,
  PullSessionHost,
  PullSessionHostCoordinator,
  normalizePullSessionInsufficientCreditError,
  serializeWorkerError,
  type PullSessionCommand,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { WorkerBridgeTransport } from '@silurus/ooxml-core';
import type { BodyElement, DocxDocumentModel } from './types.js';

export interface DocxDocumentCursorArchive {
  open_document_cursor(operationId: number, generation: number): void;
  pull_document_chunk(
    sequence: number,
    operationId: number,
    generation: number,
    byteCredit: number,
  ): Uint8Array;
  document_chunk_done(): boolean;
  /** `undefined` when the package has no document-cursor checkpoint. */
  document_cursor_resource_usage?(): Uint8Array | undefined;
  acknowledge_document_chunk(
    sequence: number,
    operationId: number,
    generation: number,
  ): void;
  cancel_document_cursor(): void;
  close_document_session(): void;
}

/**
 * Destructive, bounded cursor over an already-materialized compatibility model.
 * This is used only when render-worker capability detection requires the model
 * to fall back to Window. Moving one body block at a time avoids recreating the
 * former whole-document JSON spike while preserving the same public model.
 */
export class MaterializedDocumentCursorArchive implements DocxDocumentCursorArchive {
  private body: Array<BodyElement | undefined>;
  private document: DocxDocumentModel | null;
  private prepared: { sequence: number; bytes: Uint8Array; done: boolean } | null = null;
  private index = 0;
  private identity: PullSessionIdentity<number> | null = null;
  private completed = false;

  constructor(document: DocxDocumentModel) {
    this.body = document.body;
    document.body = [];
    this.document = document;
  }

  open_document_cursor(operationId: number, generation: number): void {
    if (this.identity) throw new Error('materialized DOCX cursor is already open');
    this.identity = { sessionId: generation, operationId, generation };
  }

  pull_document_chunk(
    sequence: number,
    operationId: number,
    generation: number,
    byteCredit: number,
  ): Uint8Array {
    this.assertIdentity(operationId, generation);
    if (this.completed) throw new Error('materialized DOCX cursor is complete');
    if (sequence !== this.index) throw new Error('materialized DOCX cursor sequence mismatch');
    if (!this.prepared) {
      const element = this.body[this.index];
      const unit = element === undefined
        ? { kind: 'complete', document: this.requireDocument() }
        : { kind: 'body', body: [element] };
      this.prepared = {
        sequence,
        bytes: new TextEncoder().encode(JSON.stringify(unit)),
        done: element === undefined,
      };
    }
    if (this.prepared.bytes.byteLength > byteCredit) {
      throw new PullSessionInsufficientCreditError({
        requiredBytes: this.prepared.bytes.byteLength,
        offeredBytes: byteCredit,
      });
    }
    // The returned buffer is transferred and detached; retain the prepared
    // bytes so an unacknowledged pull can be retried without reserialization.
    return this.prepared.bytes.slice();
  }

  document_chunk_done(): boolean {
    if (!this.prepared) throw new Error('materialized DOCX cursor has no prepared unit');
    return this.prepared.done;
  }

  acknowledge_document_chunk(
    sequence: number,
    operationId: number,
    generation: number,
  ): void {
    this.assertIdentity(operationId, generation);
    if (!this.prepared || sequence !== this.prepared.sequence || sequence !== this.index) {
      throw new Error('materialized DOCX acknowledgement sequence mismatch');
    }
    const done = this.prepared.done;
    this.prepared = null;
    if (done) {
      this.completed = true;
      this.body = [];
      this.document = null;
      return;
    }
    this.body[this.index] = undefined;
    this.index += 1;
  }

  cancel_document_cursor(): void {
    this.prepared = null;
    this.body = [];
    this.document = null;
    this.completed = true;
  }

  close_document_session(): void {
    this.cancel_document_cursor();
  }

  private assertIdentity(operationId: number, generation: number): void {
    if (
      !this.identity
      || this.identity.operationId !== operationId
      || this.identity.generation !== generation
    ) throw new Error('materialized DOCX cursor identity mismatch');
  }

  private requireDocument(): DocxDocumentModel {
    if (!this.document) throw new Error('materialized DOCX terminal document is unavailable');
    return this.document;
  }
}

const MAX_DOCUMENT_UNIT_BYTES = Math.max(
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
);

/** DOCX-owned adapter from the sequential Rust body cursor to the common
 * correlated pull/ACK protocol. It owns no semantic parsing or layout state. */
export class DocumentPullWorker {
  private coordinator = new PullSessionHostCoordinator();
  private host: PullSessionHost<ArrayBuffer, number> | null = null;
  private identity: PullSessionIdentity<number> | null = null;

  constructor(
    private readonly archive: () => DocxDocumentCursorArchive | null | undefined,
    private readonly executeArchive: <T>(
      operation: (archive: DocxDocumentCursorArchive) => T,
    ) => T = (operation) => operation(this.requireArchive()),
  ) {}

  open(identity: PullSessionIdentity<number>): void {
    if (this.host) throw new Error('a DOCX document pull session is already active');
    this.executeArchive((archive) => {
      archive.open_document_cursor(identity.operationId, identity.generation);
    });
    let sequence = 0;
    this.identity = identity;
    this.host = new PullSessionHost<ArrayBuffer, number>({
      ...identity,
      maxByteCredit: MAX_DOCUMENT_UNIT_BYTES,
      coordinator: this.coordinator,
      driver: {
        pull: (byteCredit) => {
          let bytes: Uint8Array;
          try {
            bytes = this.executeArchive((archive) => archive.pull_document_chunk(
              sequence,
              identity.operationId,
              identity.generation,
              byteCredit,
            ));
          } catch (error) {
            const insufficient = normalizePullSessionInsufficientCreditError(
              error,
              byteCredit,
              MAX_DOCUMENT_UNIT_BYTES,
            );
            if (insufficient) throw insufficient;
            throw error;
          }
          const payload = exactTransferableArrayBuffer(bytes);
          return {
            payload,
            byteLength: payload.byteLength,
            done: this.executeArchive((archive) => archive.document_chunk_done()),
            transfer: [payload],
          };
        },
        measureChunk: ({ payload }) => payload.byteLength,
        acknowledge: (acknowledgedSequence) => {
          if (acknowledgedSequence !== sequence) {
            throw new Error('DOCX document acknowledgement sequence mismatch');
          }
          this.executeArchive((archive) => archive.acknowledge_document_chunk(
            sequence,
            identity.operationId,
            identity.generation,
          ));
          sequence += 1;
        },
        cancel: () => this.executeArchive((archive) => archive.cancel_document_cursor()),
        close: () => this.executeArchive((archive) => archive.close_document_session()),
        resourceUsage: () => {
          const bytes = this.executeArchive((archive) =>
            archive.document_cursor_resource_usage?.());
          return bytes ? decodeOoxmlResourceUsage(bytes) : undefined;
        },
      },
    });
  }

  dispatch(
    command: PullSessionCommand<number>,
    post: (response: PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) => void,
  ): Promise<void> {
    if (!this.host || !this.identity) {
      post({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'error',
        sessionId: command.sessionId,
        operationId: command.operationId,
        generation: command.generation,
        requestId: command.requestId,
        error: serializeWorkerError(new Error('DOCX document pull session is not open')),
      });
      return Promise.resolve();
    }
    return this.host.dispatch(command, post);
  }

  async reset(): Promise<void> {
    if (this.host) {
      try {
        if (this.archive()) {
          this.executeArchive((archive) => archive.close_document_session());
        }
      } finally {
        this.host = null;
        this.identity = null;
        this.coordinator = new PullSessionHostCoordinator();
      }
    }
  }

  private requireArchive(): DocxDocumentCursorArchive {
    const archive = this.archive();
    if (!archive) throw new Error('No docx loaded');
    return archive;
  }
}

export function isDocumentPullCommand(value: unknown): value is PullSessionCommand<number> {
  return !!value && typeof value === 'object'
    && (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

/** In-realm transport used by the render worker and Node adapters. It exercises
 * the identical host state machine without posting through a Worker boundary. */
export function createLocalDocumentPullTransport(
  worker: DocumentPullWorker,
): WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>> {
  let requestId = 1;
  return {
    request(build) {
      const command = build(requestId++) as PullSessionCommand<number>;
      return new Promise((resolve, reject) => {
        void worker.dispatch(command, resolve).catch(reject);
      });
    },
    forgetOrphaned() {},
    terminate() {},
  };
}
