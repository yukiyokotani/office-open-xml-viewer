// Correlated pull protocol used by the DOCX, PPTX and XLSX pull clients and
// workers: request, session, operation and generation identities; pull with a
// byte credit, ack, lease release, cancel and close commands; and optional
// usage snapshots on responses. Format payloads stay opaque (`TPayload`);
// there is deliberately no common page/slide/row chunk type, because unit
// boundaries and readiness barriers belong to each format.
import type { WorkerBridgeTransport, WorkerRequestOptions } from './bridge.js';
import type { OoxmlResourceUsageSnapshot } from '../errors/ooxml-error.js';
import {
  deserializeWorkerError,
  serializeWorkerError,
  type WorkerErrorPayload,
} from './error-wire.js';
import {
  isPullSessionInsufficientCreditError,
} from './pull-credit-error.js';
export { PULL_SESSION_INSUFFICIENT_CREDIT_CODE } from './pull-credit-error.js';

const DEFAULT_PULL_CANCEL_GRACE_MS = 1_000;
export const PULL_SESSION_PROTOCOL = 'ooxml-pull-v1' as const;

export type PullSessionKey = string | number;

export interface PullSessionIdentity<TSessionId extends PullSessionKey = PullSessionKey> {
  readonly sessionId: TSessionId;
  readonly operationId: number;
  readonly generation: number;
}

interface PullCommandBase<TSessionId extends PullSessionKey> extends PullSessionIdentity<TSessionId> {
  readonly protocol: typeof PULL_SESSION_PROTOCOL;
  readonly requestId: number;
}

export type PullSessionCommand<TSessionId extends PullSessionKey = PullSessionKey> =
  | (PullCommandBase<TSessionId> & {
      readonly kind: 'pull';
      readonly sequence: number;
      readonly byteCredit: number;
    })
  | (PullCommandBase<TSessionId> & { readonly kind: 'ack'; readonly sequence: number })
  | (PullCommandBase<TSessionId> & { readonly kind: 'release'; readonly leaseId: number })
  | (PullCommandBase<TSessionId> & { readonly kind: 'cancel'; readonly reason: PullCancelReason })
  | (PullCommandBase<TSessionId> & { readonly kind: 'close' });

interface PullResponseBase<TSessionId extends PullSessionKey> extends PullSessionIdentity<TSessionId> {
  readonly protocol: typeof PULL_SESSION_PROTOCOL;
  readonly requestId: number;
  readonly usage?: OoxmlResourceUsageSnapshot;
}

export type PullSessionResponse<TPayload, TSessionId extends PullSessionKey = PullSessionKey> =
  | (PullResponseBase<TSessionId> & {
      readonly kind: 'chunk';
      readonly sequence: number;
      readonly byteLength: number;
      readonly done: boolean;
      readonly payload: TPayload;
      readonly leaseId?: number;
    })
  | (PullResponseBase<TSessionId> & {
      readonly kind: 'accepted';
      readonly command: 'ack' | 'release' | 'cancel' | 'close';
    })
  | (PullResponseBase<TSessionId> & { readonly kind: 'error'; readonly error: WorkerErrorPayload });

export type PullCancelReason = 'closed' | 'timeout' | 'abort' | 'protocol-error' | 'request-error';
export type PullRequestOptions = Omit<WorkerRequestOptions, 'onOrphanedResponse'>;

export interface PullSessionClientOptions<
  TPayload,
  TSessionId extends PullSessionKey = PullSessionKey,
> extends PullSessionIdentity<TSessionId> {
  readonly maxByteCredit: number;
  readonly timeoutMs?: number;
  readonly cancelGraceMs?: number;
  /** Releases only main-thread ownership of transferred payload resources. */
  readonly disposeTransferred?: (payload: TPayload) => void;
}

export interface PullChunk<TPayload> {
  readonly done: boolean;
  readonly sequence: number;
  readonly byteLength: number;
  readonly payload: TPayload;
  readonly leaseId?: number;
  readonly usage?: OoxmlResourceUsageSnapshot;
  /** Local-only cleanup of transfer-backed resources such as ImageBitmap. */
  disposeTransferred(): void;
  /** Correlated worker acceptance. The next pull is blocked until this resolves. */
  ack(options?: PullRequestOptions): Promise<void>;
  /** Correlated release of an optional worker-retained numeric lease. */
  release(options?: PullRequestOptions): Promise<void>;
}

function validatePositiveInteger(value: number, label: string): void {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(`${label} must be a positive safe integer`);
  }
}

function validateSessionId(value: PullSessionKey): void {
  if (
    (typeof value === 'string' && value.length > 0) ||
    (typeof value === 'number' && Number.isSafeInteger(value) && value > 0)
  ) return;
  throw new RangeError('session id must be a non-empty string or positive safe integer');
}

/** Main-thread half of the shared, format-neutral bounded pull protocol. */
export class BoundedPullSession<
  TPayload,
  TSessionId extends PullSessionKey = PullSessionKey,
> {
  private readonly bridge: WorkerBridgeTransport<PullSessionResponse<TPayload, TSessionId>>;
  private readonly options: PullSessionClientOptions<TPayload, TSessionId>;
  private sequence = 0;
  private pulling = false;
  private outstanding = false;
  private ending = false;
  private completed = false;
  private lifecyclePromise?: Promise<void>;
  private usage?: OoxmlResourceUsageSnapshot;
  private readonly orphanedRequestIds = new Set<number>();
  private readonly transferDisposers = new Set<() => void>();

  constructor(
    bridge: WorkerBridgeTransport<PullSessionResponse<TPayload, TSessionId>>,
    options: PullSessionClientOptions<TPayload, TSessionId>,
  ) {
    validateSessionId(options.sessionId);
    validatePositiveInteger(options.operationId, 'operation id');
    validatePositiveInteger(options.generation, 'generation');
    validatePositiveInteger(options.maxByteCredit, 'max byte credit');
    if (options.cancelGraceMs !== undefined) {
      validatePositiveInteger(options.cancelGraceMs, 'cancel grace');
    }
    this.bridge = bridge;
    this.options = options;
  }

  async pull(byteCredit: number, requestOptions?: PullRequestOptions): Promise<PullChunk<TPayload>> {
    this.validateCredit(byteCredit);
    if (this.ending || this.completed) throw new Error('pull session is closed');
    if (this.pulling) throw new Error('a pull request is already in flight');
    if (this.outstanding) throw new Error('acknowledge the current chunk before the next pull');
    this.pulling = true;
    const expectedSequence = this.sequence;
    try {
      const response = await this.request(
        (requestId) => ({
          protocol: PULL_SESSION_PROTOCOL,
          kind: 'pull',
          ...this.identity(),
          requestId,
          sequence: expectedSequence,
          byteCredit,
        }),
        requestOptions,
      );
      this.validateResponseIdentity(response);
      if (response.usage) this.usage = response.usage;
      if (response.kind === 'error') throw deserializeWorkerError(response.error);
      if (response.kind !== 'chunk') throw new Error(`expected chunk response, received ${response.kind}`);
      if (this.ending) {
        this.disposePayload(response.payload);
        throw new Error('pull session is closed');
      }
      try {
        this.validateChunk(response, expectedSequence, byteCredit);
      } catch (error) {
        this.disposePayload(response.payload);
        throw error;
      }
      this.outstanding = true;
      return this.makeChunk(response);
    } catch (error) {
      // An indivisible format unit may be retried with more credit using the
      // same sequence and operation identity. Every other failure converges on
      // cancellation because its ownership state is not known to be reusable.
      if (!isPullSessionInsufficientCreditError(error) && !this.ending && !this.completed) {
        void this.cancel('protocol-error').catch(() => undefined);
      }
      throw error;
    } finally {
      this.pulling = false;
    }
  }

  cancel(reason: PullCancelReason = 'closed', options?: PullRequestOptions): Promise<void> {
    if (this.lifecyclePromise) return this.lifecyclePromise;
    this.ending = true;
    this.disposeAllTransferred();
    this.lifecyclePromise = this.lifecycleControl(
      this.control('cancel', options, (requestId) => ({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'cancel',
        ...this.identity(),
        requestId,
        reason,
      })),
    );
    return this.lifecyclePromise;
  }

  close(options?: PullRequestOptions): Promise<void> {
    if (this.lifecyclePromise) return this.lifecyclePromise;
    this.ending = true;
    this.disposeAllTransferred();
    this.lifecyclePromise = this.lifecycleControl(
      this.control('close', options, (requestId) => ({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'close',
        ...this.identity(),
        requestId,
      })),
    );
    return this.lifecyclePromise;
  }

  get usageCheckpoint(): OoxmlResourceUsageSnapshot | undefined {
    return this.usage;
  }

  private makeChunk(
    response: Extract<PullSessionResponse<TPayload, TSessionId>, { kind: 'chunk' }>,
  ): PullChunk<TPayload> {
    let disposed = false;
    let ackPromise: Promise<void> | undefined;
    let releasePromise: Promise<void> | undefined;
    const disposeTransferred = (): void => {
      if (disposed) return;
      disposed = true;
      this.transferDisposers.delete(disposeTransferred);
      this.disposePayload(response.payload);
    };
    this.transferDisposers.add(disposeTransferred);

    return {
      done: response.done,
      sequence: response.sequence,
      byteLength: response.byteLength,
      payload: response.payload,
      leaseId: response.leaseId,
      usage: response.usage,
      disposeTransferred,
      ack: (options) => {
        if (ackPromise) return ackPromise;
        if (this.ending) return Promise.resolve();
        ackPromise = this.control('ack', options, (requestId) => ({
          protocol: PULL_SESSION_PROTOCOL,
          kind: 'ack',
          ...this.identity(),
          requestId,
          sequence: response.sequence,
        }))
          .then(() => {
            this.outstanding = false;
            this.sequence++;
            if (response.done) this.completed = true;
          })
          .catch((error: unknown) => {
            if (!this.ending) void this.cancel('request-error').catch(() => undefined);
            throw error;
          });
        return ackPromise;
      },
      release: (options) => {
        if (releasePromise) return releasePromise;
        if (response.leaseId === undefined || this.ending) return Promise.resolve();
        releasePromise = this.control('release', options, (requestId) => ({
          protocol: PULL_SESSION_PROTOCOL,
          kind: 'release',
          ...this.identity(),
          requestId,
          leaseId: response.leaseId as number,
        })).catch((error: unknown) => {
          if (!this.ending) void this.cancel('request-error').catch(() => undefined);
          throw error;
        });
        return releasePromise;
      },
    };
  }

  private async control(
    expected: 'ack' | 'release' | 'cancel' | 'close',
    options: PullRequestOptions | undefined,
    build: (requestId: number) => PullSessionCommand<TSessionId>,
  ): Promise<void> {
    const response = await this.request(
      build,
      options,
      expected !== 'cancel' && expected !== 'close',
    );
    this.validateResponseIdentity(response);
    if (response.usage) this.usage = response.usage;
    if (response.kind === 'error') {
      if (
        response.error.code === 'ooxml-stale-lifecycle' &&
        (expected === 'cancel' || expected === 'close')
      ) {
        this.forgetLifecycleOrphans();
        return;
      }
      throw deserializeWorkerError(response.error);
    }
    if (response.kind !== 'accepted' || response.command !== expected) {
      throw new Error(`expected ${expected} acceptance`);
    }
    if (expected === 'cancel' || expected === 'close') {
      // One Worker preserves postMessage ordering: once the package-wide host
      // queue accepts lifecycle cleanup, any earlier pull response was already
      // delivered/disposed, so its tombstone can no longer own a future value.
      this.forgetLifecycleOrphans();
    }
  }

  private request(
    build: (requestId: number) => PullSessionCommand<TSessionId>,
    options?: PullRequestOptions,
    useOrdinaryTimeout = true,
  ): Promise<PullSessionResponse<TPayload, TSessionId>> {
    return this.bridge.request(build, undefined, {
      timeoutMs: useOrdinaryTimeout ? (options?.timeoutMs ?? this.options.timeoutMs) : false,
      signal: options?.signal,
      onOrphanedResponse: (response) => {
        if (response.kind === 'chunk') this.disposePayload(response.payload);
      },
      onCancel: (id, reason) => {
        this.orphanedRequestIds.add(id);
        try {
          options?.onCancel?.(id, reason);
        } finally {
          if (!this.ending && !this.completed) void this.cancel(reason).catch(() => undefined);
        }
      },
    });
  }

  private identity(): PullSessionIdentity<TSessionId> {
    return {
      sessionId: this.options.sessionId,
      operationId: this.options.operationId,
      generation: this.options.generation,
    };
  }

  private validateResponseIdentity(response: PullSessionResponse<TPayload, TSessionId>): void {
    if (
      response.protocol !== PULL_SESSION_PROTOCOL ||
      response.sessionId !== this.options.sessionId ||
      response.operationId !== this.options.operationId ||
      response.generation !== this.options.generation
    ) {
      if (response.kind === 'chunk') this.disposePayload(response.payload);
      throw new Error('stale or mismatched pull session response');
    }
  }

  private validateCredit(byteCredit: number): void {
    validatePositiveInteger(byteCredit, 'byte credit');
    if (byteCredit > this.options.maxByteCredit) {
      throw new RangeError(`byte credit exceeds session maximum ${this.options.maxByteCredit}`);
    }
  }

  private validateChunk(
    response: Extract<PullSessionResponse<TPayload, TSessionId>, { kind: 'chunk' }>,
    sequence: number,
    byteCredit: number,
  ): void {
    if (response.sequence !== sequence) throw new Error('pull response sequence mismatch');
    if (!Number.isSafeInteger(response.byteLength) || response.byteLength < 0) {
      throw new RangeError('chunk byte length must be a non-negative safe integer');
    }
    if (response.byteLength > byteCredit) throw new RangeError('chunk exceeds byte credit');
    if (response.leaseId !== undefined) validatePositiveInteger(response.leaseId, 'lease id');
  }

  private disposeAllTransferred(): void {
    for (const dispose of [...this.transferDisposers]) dispose();
  }

  private lifecycleControl(control: Promise<void>): Promise<void> {
    const graceMs = this.options.cancelGraceMs ?? DEFAULT_PULL_CANCEL_GRACE_MS;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const fallback = new Promise<void>((_resolve, reject) => {
      timer = setTimeout(() => {
        this.bridge.terminate();
        reject(new Error(`worker did not accept lifecycle command within ${graceMs}ms`));
      }, graceMs);
    });
    return Promise.race([control, fallback])
      .catch((error: unknown) => {
        this.bridge.terminate();
        throw error;
      })
      .finally(() => {
        if (timer !== undefined) clearTimeout(timer);
      });
  }

  private forgetLifecycleOrphans(): void {
    this.bridge.forgetOrphaned(this.orphanedRequestIds);
    this.orphanedRequestIds.clear();
  }

  private disposePayload(payload: TPayload): void {
    try {
      this.options.disposeTransferred?.(payload);
    } catch {
      // Transfer cleanup is best-effort and independent of worker lease state.
    }
  }
}

export interface PullSessionHostChunk<TPayload> {
  readonly payload: TPayload;
  /** Driver estimate is diagnostic only; host credit uses measureChunk(). */
  readonly byteLength: number;
  readonly done: boolean;
  /** Worker-retained resource identity, distinct from the transferred payload. */
  readonly leaseId?: number;
  /** Bytes retained in the worker under leaseId; required whenever leaseId exists. */
  readonly retainedBytes?: number;
  /** Ownership passes to postMessage only after host validation succeeds. */
  readonly transfer?: Transferable[];
}

export interface PullSessionHostDriver<TPayload> {
  pull(byteCredit: number): PullSessionHostChunk<TPayload> | Promise<PullSessionHostChunk<TPayload>>;
  /**
   * Derive actual payload/transfer bytes without trusting pull() metadata.
   * The host independently counts ArrayBuffer transfer bytes. This adapter must
   * additionally measure non-ArrayBuffer ownership such as SharedArrayBuffer,
   * MessagePort-backed data, ImageBitmap, and semantic payload allocations.
   */
  measureChunk(chunk: PullSessionHostChunk<TPayload>): number;
  acknowledge?(sequence: number): void | Promise<void>;
  releaseLease?(leaseId: number): void | Promise<void>;
  /** Dispose transfer-backed objects when a driver chunk is rejected before post. */
  disposeInvalidChunk?(chunk: PullSessionHostChunk<TPayload>): void | Promise<void>;
  cancel?(): void | Promise<void>;
  close?(): void | Promise<void>;
  resourceUsage?(): OoxmlResourceUsageSnapshot | undefined;
}

export interface PullSessionHostOptions<
  TPayload,
  TSessionId extends PullSessionKey = PullSessionKey,
> extends PullSessionIdentity<TSessionId> {
  readonly maxByteCredit: number;
  readonly driver: PullSessionHostDriver<TPayload>;
  /** Share one coordinator across every operation backed by a package session. */
  readonly coordinator: PullSessionHostCoordinator;
  /** Test/recovery injection; ordinary generations start wire lease IDs at 1. */
  readonly wireLeaseIdStart?: number;
}

export interface PullSessionPost<TPayload, TSessionId extends PullSessionKey = PullSessionKey> {
  readonly response: PullSessionResponse<TPayload, TSessionId>;
  readonly transfer?: Transferable[];
}

/** Session-level gate: at most one operation may own a produced, unacked chunk. */
export class PullSessionHostCoordinator {
  private owner?: symbol;
  private queue: Promise<void> = Promise.resolve();
  private readonly leases = new Map<symbol, Map<number, number>>();
  private retainedBytes = 0;
  private retainedCount = 0;
  private readonly maxRetainedBytes: number;
  private readonly maxRetainedCount: number;
  private readonly cleanups = new Set<() => Promise<void>>();
  private readonly pendingFatalCleanups: Array<() => Promise<void>> = [];
  private poisonRunning = false;
  private fatal?: WorkerErrorPayload;

  constructor(options?: { maxRetainedBytes?: number; maxRetainedCount?: number }) {
    this.maxRetainedBytes = options?.maxRetainedBytes ?? 64 * 1024 * 1024;
    this.maxRetainedCount = options?.maxRetainedCount ?? 256;
    validatePositiveInteger(this.maxRetainedBytes, 'max retained lease bytes');
    validatePositiveInteger(this.maxRetainedCount, 'max retained lease count');
  }

  enqueue<T>(run: () => Promise<T>): Promise<T> {
    const result = this.queue.then(run, run);
    this.queue = result.then(() => undefined, () => undefined);
    return result;
  }

  acquire(owner: symbol): boolean {
    if (this.owner !== undefined) return this.owner === owner;
    this.owner = owner;
    return true;
  }

  release(owner: symbol): void {
    if (this.owner === owner) this.owner = undefined;
  }

  retainLease(owner: symbol, leaseId: number, bytes: number): void {
    if (!Number.isSafeInteger(bytes) || bytes < 0) throw new RangeError('retained lease bytes are invalid');
    const owned = this.leases.get(owner) ?? new Map<number, number>();
    if (owned.has(leaseId)) throw new Error('driver returned a duplicate lease id');
    if (this.retainedCount + 1 > this.maxRetainedCount) throw new RangeError('retained lease count exceeds limit');
    if (this.retainedBytes + bytes > this.maxRetainedBytes) throw new RangeError('retained lease bytes exceed limit');
    owned.set(leaseId, bytes);
    this.leases.set(owner, owned);
    this.retainedCount++;
    this.retainedBytes += bytes;
  }

  releaseLease(owner: symbol, leaseId: number): void {
    const owned = this.leases.get(owner);
    const bytes = owned?.get(leaseId);
    if (bytes === undefined) return;
    owned?.delete(leaseId);
    if (owned?.size === 0) this.leases.delete(owner);
    this.retainedCount--;
    this.retainedBytes -= bytes;
  }

  registerCleanup(cleanup: () => Promise<void>): () => void {
    if (this.fatal) {
      if (this.poisonRunning) this.pendingFatalCleanups.push(cleanup);
      else void this.enqueue(cleanup).catch(() => undefined);
      return () => undefined;
    }
    this.cleanups.add(cleanup);
    return () => this.cleanups.delete(cleanup);
  }

  get fatalError(): WorkerErrorPayload | undefined {
    return this.fatal;
  }

  get registeredHostCount(): number {
    return this.cleanups.size;
  }

  async poison(error: WorkerErrorPayload): Promise<WorkerErrorPayload> {
    this.fatal ??= error;
    if (this.poisonRunning) return this.fatal;
    this.poisonRunning = true;
    this.pendingFatalCleanups.push(...this.cleanups);
    try {
      let cleanup: (() => Promise<void>) | undefined;
      while ((cleanup = this.pendingFatalCleanups.shift()) !== undefined) {
        await cleanup().catch(() => undefined);
      }
    } finally {
      this.poisonRunning = false;
    }
    return this.fatal;
  }
}

/** Worker-side serialized command state machine for one operation generation. */
export class PullSessionHost<
  TPayload,
  TSessionId extends PullSessionKey = PullSessionKey,
> {
  private readonly options: PullSessionHostOptions<TPayload, TSessionId>;
  private readonly coordinator: PullSessionHostCoordinator;
  private readonly coordinatorOwner = Symbol('pull-session-host');
  private readonly unregisterCleanup: () => void;
  private sequence = 0;
  private unacked?: { sequence: number; done: boolean };
  private readonly leases = new Map<
    number,
    { readonly driverLeaseId: number; readonly retainedBytes: number }
  >();
  private readonly activeDriverLeases = new Set<number>();
  private nextWireLeaseId: number;
  private cancelRequested = false;
  private cancelComplete = false;
  private closeRequested = false;
  private closeComplete = false;
  private driverCancelComplete = false;
  private driverCloseComplete = false;
  private completed = false;

  constructor(options: PullSessionHostOptions<TPayload, TSessionId>) {
    validateSessionId(options.sessionId);
    validatePositiveInteger(options.operationId, 'operation id');
    validatePositiveInteger(options.generation, 'generation');
    validatePositiveInteger(options.maxByteCredit, 'max byte credit');
    if (options.wireLeaseIdStart !== undefined) {
      validatePositiveInteger(options.wireLeaseIdStart, 'wire lease id start');
    }
    this.options = options;
    this.coordinator = options.coordinator;
    this.nextWireLeaseId = options.wireLeaseIdStart ?? 1;
    this.unregisterCleanup = this.coordinator.registerCleanup(() => this.forceFatalCleanup());
  }

  /**
   * Execute and post one command as a single ownership transaction. A throwing
   * post (including structured-clone failure) rolls back worker-owned payload,
   * lease, coordinator, and operation state before rejecting.
   */
  dispatch(
    command: PullSessionCommand<TSessionId>,
    post: (
      response: PullSessionResponse<TPayload, TSessionId>,
      transfer?: Transferable[],
    ) => void,
  ): Promise<void> {
    const run = async (): Promise<void> => {
      const output = await this.execute(command);
      try {
        post(output.response, output.transfer);
      } catch (error) {
        await this.rollbackFailedPost(output);
        throw error;
      }
    };
    return this.coordinator.enqueue(run);
  }

  private async rollbackFailedPost(output: PullSessionPost<TPayload, TSessionId>): Promise<void> {
    const response = output.response;
    if (response.kind === 'chunk') {
      const lease = response.leaseId === undefined ? undefined : this.leases.get(response.leaseId);
      try {
        await this.options.driver.disposeInvalidChunk?.({
          payload: response.payload,
          byteLength: response.byteLength,
          done: response.done,
          leaseId: lease?.driverLeaseId,
          retainedBytes: lease?.retainedBytes,
          transfer: output.transfer,
        });
      } catch {
        // Continue through every ownership boundary after a failed post.
      }
    }
    this.unacked = undefined;
    this.coordinator.release(this.coordinatorOwner);
    for (const [wireLeaseId, lease] of [...this.leases]) {
      try {
        await this.options.driver.releaseLease?.(lease.driverLeaseId);
      } catch {
        // The host must forget an unpostable lease even if driver cleanup fails.
      } finally {
        this.leases.delete(wireLeaseId);
        this.activeDriverLeases.delete(lease.driverLeaseId);
        this.coordinator.releaseLease(this.coordinatorOwner, wireLeaseId);
      }
    }
    this.cancelRequested = true;
    if (!this.driverCancelComplete) {
      try {
        await this.options.driver.cancel?.();
        this.driverCancelComplete = true;
      } catch {
        // Posting already failed; cleanup remains best-effort and idempotent.
      }
    }
    this.unregisterCleanup();
  }

  private async execute(
    command: PullSessionCommand<TSessionId>,
  ): Promise<PullSessionPost<TPayload, TSessionId>> {
    try {
      if (this.isStaleLifecycle(command)) {
        const lifecycle = command.kind === 'cancel' ? 'cancel' : 'close';
        return this.sameOperationIdentity(command)
          ? { response: this.accepted(command, lifecycle, true) }
          : {
              response: this.errorResponse(command, {
                message: 'stale lifecycle targets another session or operation',
                errorName: 'PullSessionProtocolError',
                code: 'ooxml-stale-lifecycle',
              }),
            };
      }
      this.validateCommandIdentity(command);
      const fatal = this.coordinator.fatalError;
      if (fatal) {
        if (command.kind === 'pull') return { response: this.errorResponse(command, fatal) };
        if (command.kind === 'cancel') await this.cancel();
        else if (command.kind === 'close') await this.close();
        else if (command.kind === 'release') await this.release(command.leaseId);
        return { response: this.accepted(command, command.kind) };
      }
      switch (command.kind) {
        case 'pull':
          return await this.pull(command);
        case 'ack':
          await this.ack(command.sequence);
          return { response: this.accepted(command, 'ack') };
        case 'release':
          await this.release(command.leaseId);
          return { response: this.accepted(command, 'release') };
        case 'cancel':
          await this.cancel();
          return { response: this.accepted(command, 'cancel') };
        case 'close':
          await this.close();
          return { response: this.accepted(command, 'close') };
      }
    } catch (error) {
      let payload = serializeWorkerError(error);
      if (payload.code === 'ooxml-resource-limit') {
        payload = await this.coordinator.poison(payload);
      }
      return {
        response: this.errorResponse(command, payload),
      };
    }
  }

  private async pull(
    command: Extract<PullSessionCommand<TSessionId>, { kind: 'pull' }>,
  ): Promise<PullSessionPost<TPayload, TSessionId>> {
    if (this.closeRequested || this.cancelRequested || this.completed) throw new Error('pull session is closed');
    if (this.unacked) throw new Error('previous chunk is not acknowledged');
    if (!Number.isSafeInteger(command.sequence) || command.sequence < 0 || command.sequence !== this.sequence) {
      throw new Error('pull command sequence mismatch');
    }
    this.validateHostCredit(command.byteCredit);
    if (!this.coordinator.acquire(this.coordinatorOwner)) {
      throw new Error('another operation has an unacknowledged package chunk');
    }
    let chunk: PullSessionHostChunk<TPayload>;
    try {
      chunk = await this.options.driver.pull(command.byteCredit);
    } catch (error) {
      this.coordinator.release(this.coordinatorOwner);
      throw error;
    }
    let addedLease = false;
    let duplicateDriverLease = false;
    let wireLeaseId: number | undefined;
    let byteLength: number;
    try {
      const measuredBytes = this.options.driver.measureChunk(chunk);
      const arrayBufferBytes = this.arrayBufferTransferBytes(chunk.transfer);
      if (measuredBytes < arrayBufferBytes) {
        throw new RangeError('measured chunk bytes are below ArrayBuffer transfer bytes');
      }
      byteLength = Math.max(measuredBytes, arrayBufferBytes);
      if (chunk.leaseId !== undefined) {
        validatePositiveInteger(chunk.leaseId, 'lease id');
        if (chunk.retainedBytes === undefined) throw new Error('retained lease bytes are required');
        if (this.activeDriverLeases.has(chunk.leaseId)) {
          duplicateDriverLease = true;
          throw new Error('driver returned an active duplicate lease id');
        }
        wireLeaseId = this.allocateWireLeaseId();
        this.coordinator.retainLease(this.coordinatorOwner, wireLeaseId, chunk.retainedBytes);
        this.leases.set(wireLeaseId, {
          driverLeaseId: chunk.leaseId,
          retainedBytes: chunk.retainedBytes,
        });
        this.activeDriverLeases.add(chunk.leaseId);
        addedLease = true;
      } else if (chunk.retainedBytes !== undefined) {
        throw new Error('retained lease bytes require a lease id');
      }
      if (!Number.isSafeInteger(byteLength) || byteLength < 0) {
        throw new RangeError('host chunk byte length must be a non-negative safe integer');
      }
      if (byteLength > command.byteCredit) throw new RangeError('host chunk exceeds byte credit');
    } catch (error) {
      let cleanupError: unknown;
      try {
        await this.options.driver.disposeInvalidChunk?.(chunk);
      } catch (failure) {
        cleanupError = failure;
      }
      if (addedLease && wireLeaseId !== undefined) {
        try {
          await this.release(wireLeaseId);
        } catch (failure) {
          cleanupError ??= failure;
        }
      } else if (chunk.leaseId !== undefined && !duplicateDriverLease) {
        try {
          await this.options.driver.releaseLease?.(chunk.leaseId);
        } catch (failure) {
          cleanupError ??= failure;
        }
      }
      if (duplicateDriverLease) {
        try {
          await this.cancel();
        } catch (failure) {
          cleanupError ??= failure;
        }
      }
      this.coordinator.release(this.coordinatorOwner);
      if (cleanupError) throw cleanupError;
      throw error;
    }
    this.unacked = { sequence: this.sequence, done: chunk.done };
    return {
      response: {
        kind: 'chunk',
        protocol: PULL_SESSION_PROTOCOL,
        ...this.identity(),
        requestId: command.requestId,
        sequence: this.sequence,
        byteLength,
        done: chunk.done,
        payload: chunk.payload,
        leaseId: wireLeaseId,
        usage: this.resourceUsage(),
      },
      transfer: chunk.transfer,
    };
  }

  private async ack(sequence: number): Promise<void> {
    if (!Number.isSafeInteger(sequence) || sequence < 0) throw new RangeError('invalid ack sequence');
    if (sequence < this.sequence) return;
    if (!this.unacked || sequence !== this.sequence) throw new Error('ack sequence mismatch');
    const final = this.unacked.done;
    await this.options.driver.acknowledge?.(sequence);
    this.unacked = undefined;
    this.coordinator.release(this.coordinatorOwner);
    this.sequence++;
    if (final) {
      this.completed = true;
      this.maybeUnregisterCompleted();
    }
  }

  private async release(wireLeaseId: number): Promise<void> {
    validatePositiveInteger(wireLeaseId, 'wire lease id');
    const lease = this.leases.get(wireLeaseId);
    if (!lease) return;
    await this.options.driver.releaseLease?.(lease.driverLeaseId);
    this.leases.delete(wireLeaseId);
    this.activeDriverLeases.delete(lease.driverLeaseId);
    this.coordinator.releaseLease(this.coordinatorOwner, wireLeaseId);
    this.maybeUnregisterCompleted();
  }

  private async cancel(): Promise<void> {
    if (this.cancelComplete) return;
    this.cancelRequested = true;
    this.unacked = undefined;
    this.coordinator.release(this.coordinatorOwner);
    let error: unknown;
    try {
      await this.releaseAllLeases();
    } catch (failure) {
      error = failure;
    }
    if (!this.driverCancelComplete) {
      try {
        await this.options.driver.cancel?.();
        this.driverCancelComplete = true;
      } catch (failure) {
        error ??= failure;
      }
    }
    if (error) throw error;
    this.cancelComplete = true;
    this.unregisterCleanup();
  }

  private async close(): Promise<void> {
    if (this.closeComplete) return;
    this.closeRequested = true;
    this.unacked = undefined;
    this.coordinator.release(this.coordinatorOwner);
    let error: unknown;
    try {
      await this.releaseAllLeases();
    } catch (failure) {
      error = failure;
    }
    if (!this.driverCloseComplete) {
      try {
        await this.options.driver.close?.();
        this.driverCloseComplete = true;
      } catch (failure) {
        error ??= failure;
      }
    }
    if (error) throw error;
    this.closeComplete = true;
    this.unregisterCleanup();
  }

  private async releaseAllLeases(): Promise<void> {
    let error: unknown;
    for (const wireLeaseId of [...this.leases.keys()]) {
      try {
        await this.release(wireLeaseId);
      } catch (failure) {
        error ??= failure;
      }
    }
    if (error) throw error;
  }

  private validateCommandIdentity(command: PullSessionCommand<TSessionId>): void {
    if (
      command.protocol !== PULL_SESSION_PROTOCOL ||
      command.sessionId !== this.options.sessionId ||
      command.operationId !== this.options.operationId ||
      command.generation !== this.options.generation ||
      !Number.isSafeInteger(command.requestId) ||
      command.requestId <= 0
    ) {
      throw new Error('stale or mismatched pull session command');
    }
  }

  private validateHostCredit(byteCredit: number): void {
    validatePositiveInteger(byteCredit, 'byte credit');
    if (byteCredit > this.options.maxByteCredit) throw new RangeError('byte credit exceeds host maximum');
  }

  private accepted(
    command: PullSessionCommand<TSessionId>,
    accepted: Extract<PullSessionResponse<TPayload, TSessionId>, { kind: 'accepted' }>['command'],
    echoCommandIdentity = false,
  ): PullSessionResponse<TPayload, TSessionId> {
    return {
      kind: 'accepted',
      protocol: PULL_SESSION_PROTOCOL,
      ...(echoCommandIdentity
        ? {
            sessionId: command.sessionId,
            operationId: command.operationId,
            generation: command.generation,
          }
        : this.identity()),
      requestId: command.requestId,
      command: accepted,
      usage: this.resourceUsage(),
    };
  }

  private identity(): PullSessionIdentity<TSessionId> {
    return {
      sessionId: this.options.sessionId,
      operationId: this.options.operationId,
      generation: this.options.generation,
    };
  }

  private isStaleLifecycle(
    command: PullSessionCommand<TSessionId>,
  ): boolean {
    return (
      (command.kind === 'cancel' || command.kind === 'close') &&
      command.protocol === PULL_SESSION_PROTOCOL &&
      Number.isSafeInteger(command.requestId) &&
      command.requestId > 0 &&
      Number.isSafeInteger(command.generation) &&
      command.generation > 0 &&
      command.generation < this.options.generation
    );
  }

  private sameOperationIdentity(command: PullSessionCommand<TSessionId>): boolean {
    return (
      command.sessionId === this.options.sessionId &&
      command.operationId === this.options.operationId
    );
  }

  private errorResponse(
    command: PullSessionCommand<TSessionId>,
    error: WorkerErrorPayload,
  ): PullSessionResponse<TPayload, TSessionId> {
    return {
      kind: 'error',
      protocol: PULL_SESSION_PROTOCOL,
      sessionId: command.sessionId,
      operationId: command.operationId,
      generation: command.generation,
      requestId: command.requestId,
      error,
      // Reporting a resourceUsage() failure must not recursively mask the
      // original protocol/resource error with the same failing checkpoint.
      usage: this.errorResourceUsage(),
    };
  }

  private async forceFatalCleanup(): Promise<void> {
    this.cancelRequested = true;
    this.unacked = undefined;
    this.coordinator.release(this.coordinatorOwner);
    let error: unknown;
    for (const wireLeaseId of [...this.leases.keys()]) {
      try {
        await this.release(wireLeaseId);
      } catch (failure) {
        error ??= failure;
      }
    }
    if (!this.driverCancelComplete) {
      try {
        await this.options.driver.cancel?.();
        this.driverCancelComplete = true;
      } catch (failure) {
        error ??= failure;
      }
    }
    if (error) throw error;
    this.unregisterCleanup();
  }

  private allocateWireLeaseId(): number {
    if (!Number.isSafeInteger(this.nextWireLeaseId) || this.nextWireLeaseId <= 0) {
      throw new RangeError('wire lease id space exhausted');
    }
    return this.nextWireLeaseId++;
  }

  private arrayBufferTransferBytes(transfer?: Transferable[]): number {
    let bytes = 0;
    for (const item of transfer ?? []) {
      if (!(item instanceof ArrayBuffer)) continue;
      bytes += item.byteLength;
      if (!Number.isSafeInteger(bytes)) throw new RangeError('ArrayBuffer transfer bytes overflow');
    }
    return bytes;
  }

  private maybeUnregisterCompleted(): void {
    if (this.completed && this.leases.size === 0) this.unregisterCleanup();
  }

  private resourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    return this.options.driver.resourceUsage?.();
  }

  private errorResourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    try {
      return this.resourceUsage();
    } catch {
      return undefined;
    }
  }
}
