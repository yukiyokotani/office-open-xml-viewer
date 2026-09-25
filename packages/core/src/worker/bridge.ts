/**
 * Request/response correlation over a Web Worker.
 *
 * Each format package (pptx/docx/xlsx) drives a WASM parser in a Worker. The
 * naive pattern — attach a one-shot `message` listener keyed only on the
 * response `type` — breaks under concurrency: with two requests in flight, the
 * first arriving response of a matching type resolves the wrong promise (or, if
 * an unrelated message arrives, the listener detaches without resolving and the
 * promise hangs forever).
 *
 * `WorkerBridge` fixes this by assigning every request a monotonic id and
 * resolving against a pending-callback map keyed by that id — the proven
 * pattern. It is wire-protocol agnostic: the discriminant field (`kind` vs
 * `type`), the error shape, and any unsolicited messages (e.g. an init `ready`
 * handshake) are described by the {@link WorkerBridgeOptions} callbacks, so all
 * three packages can share one correlation mechanism without standardizing
 * their message envelopes.
 *
 * A correlated response is the happy path, but a worker can also fail *without*
 * ever posting one — a wedged parse loop, an uncaught script exception, a
 * structured-clone failure on the way back. Those leave requests hanging
 * forever. Three opt-in escape hatches settle them: a per-request `timeoutMs`,
 * an `AbortSignal`, and the worker's own `error` / `messageerror` events (which
 * reject *every* pending request, since the worker is presumed unusable). All
 * timers and listeners are torn down on settle so nothing leaks.
 */

/** The subset of the DOM `Worker` interface the bridge depends on. Keeping it
 *  structural lets tests substitute an in-memory fake. The real `Worker`
 *  satisfies this — it exposes the `error` / `messageerror` events too. */
export interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
  addEventListener(type: 'messageerror', listener: (e: MessageEvent) => void): void;
  addEventListener(type: 'error', listener: (e: ErrorEvent) => void): void;
  removeEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
  removeEventListener(type: 'messageerror', listener: (e: MessageEvent) => void): void;
  removeEventListener(type: 'error', listener: (e: ErrorEvent) => void): void;
  terminate(): void;
}

export interface WorkerBridgeOptions<TRes> {
  /**
   * Extract the correlation id from a response. Return `undefined` for
   * unsolicited messages that do not answer a request (e.g. a `ready`
   * handshake); those are routed to {@link onUnsolicited} instead.
   */
  readonly correlate: (response: TRes) => number | undefined;
  /**
   * Extract an error message or reconstructed Error from a response, or
   * `undefined` when the response is a success.
   */
  readonly toError?: (response: TRes) => string | Error | undefined;
  /** Called for every message that does not correlate to a pending request. */
  readonly onUnsolicited?: (response: TRes) => void;
  /**
   * Best-effort cleanup for a correlated response whose request is no longer
   * pending (for example, because it timed out or was aborted). Transferable
   * payloads such as ImageBitmap need an explicit owner even on this stale path.
   */
  readonly onOrphanedResponse?: (response: TRes) => void;
  /**
   * Default per-request timeout in milliseconds. When set, every {@link
   * WorkerBridge.request} that does not override it rejects if the worker has
   * not answered within this window. Omit it (the default) to wait forever —
   * the historical behaviour. A per-request `timeoutMs` takes precedence.
   */
  readonly timeoutMs?: number;
}

/** Per-request escape-hatch options for {@link WorkerBridge.request}. */
export interface WorkerRequestOptions<TResponse = unknown> {
  /**
   * Reject this request if the worker has not answered within `timeoutMs`
   * milliseconds. Overrides the bridge-level default. Omit both to wait forever.
   */
  timeoutMs?: number | false;
  /**
   * Reject this request when the signal aborts (and immediately if it is
   * already aborted). The rejection is an `Error` whose `name` is
   * `'AbortError'`.
   */
  signal?: AbortSignal;
  /** Called when timeout or abort removes the request before a response. */
  onCancel?: (id: number, reason: 'timeout' | 'abort') => void;
  /** Owns a response that arrives after this request timed out or aborted. */
  onOrphanedResponse?: (response: TResponse) => void;
}

/**
 * Correlated request surface consumed by higher-level protocols. A view keeps
 * one owning {@link WorkerBridge} (and therefore one Worker message listener)
 * while allowing a protocol such as bounded pull to describe only the response
 * variant it handles.
 */
export interface WorkerBridgeTransport<TRes> {
  request(
    build: (id: number) => unknown,
    transfer?: Transferable[],
    opts?: WorkerRequestOptions<TRes>,
  ): Promise<TRes>;
  forgetOrphaned(ids: Iterable<number>): void;
  terminate(): void;
}

interface PendingEntry<TRes> {
  resolve: (r: TRes) => void;
  reject: (e: Error) => void;
  /** Tear down this request's timer and abort listener. Idempotent; called
   *  exactly once, on whichever path settles the request first. */
  cleanup: () => void;
  onOrphanedResponse?: (response: TRes) => void;
}

/** Build an `AbortError`-flavoured Error without depending on the environment's
 *  `DOMException` constructor (absent in some worker/test runtimes). */
function abortError(): Error {
  const err = new Error('worker request aborted');
  err.name = 'AbortError';
  return err;
}

export class WorkerBridge<TRes = unknown> {
  private readonly _worker: WorkerLike;
  private readonly _opts: WorkerBridgeOptions<TRes>;
  private readonly _pending = new Map<number, PendingEntry<TRes>>();
  private readonly _orphaned = new Map<number, (response: TRes) => void>();
  private _nextId = 1;
  private _terminated = false;
  private _failure?: Error;

  constructor(worker: WorkerLike, opts: WorkerBridgeOptions<TRes>) {
    this._worker = worker;
    this._opts = opts;
    this._worker.addEventListener('message', this._handle);
    this._worker.addEventListener('messageerror', this._handleWorkerError);
    this._worker.addEventListener('error', this._handleWorkerError);
  }

  private _handle = (e: MessageEvent<TRes>): void => {
    const res = e.data;
    const id = this._opts.correlate(res);
    if (id === undefined) {
      this._opts.onUnsolicited?.(res);
      return;
    }
    const cb = this._pending.get(id);
    if (!cb) {
      try {
        const dispose = this._orphaned.get(id);
        if (dispose) {
          this._orphaned.delete(id);
          dispose(res);
        } else {
          this._opts.onOrphanedResponse?.(res);
        }
      } catch {
        // Cleanup is best-effort and must not poison unrelated requests.
      }
      return;
    }
    this._pending.delete(id);
    cb.cleanup();
    const err = this._opts.toError?.(res);
    if (err !== undefined) cb.reject(err instanceof Error ? err : new Error(err));
    else cb.resolve(res);
  };

  /**
   * The worker itself failed (uncaught script exception, load failure, or a
   * response that could not be deserialized). No correlated reply is coming, so
   * every in-flight request is rejected — the worker is presumed unusable.
   */
  private _handleWorkerError = (e: ErrorEvent | MessageEvent): void => {
    const detail = 'message' in e && e.message ? `: ${e.message}` : '';
    // HTML sandboxing gives frames without allow-same-origin an opaque origin.
    // Chromium reports a blocked Blob module-worker load there as an `error`
    // with an empty message. Both rendering modes intentionally retain module
    // Workers; opaque-origin frames are outside the supported browser setup,
    // rather than a trigger for a classic-Worker or synchronous-parse fallback.
    // Keep the normal error text when the browser supplies it, and treat this
    // hint as diagnostic rather than a classified cause.
    const opaqueOriginHint = e.type === 'error' && !detail && globalThis.origin === 'null'
      ? ': this page has an opaque origin; a sandboxed iframe without allow-same-origin can block Worker loading'
      : '';
    this._failure ??= new Error(`Worker error${detail}${opaqueOriginHint}`);
    this._rejectAll(this._failure);
  };

  /** Reject and drain every pending request, tearing down each one's timer and
   *  abort listener. Shared by worker-error handling and {@link terminate}. */
  private _rejectAll(error: Error): void {
    const entries = [...this._pending.values()];
    this._pending.clear();
    for (const cb of entries) {
      cb.cleanup();
      cb.reject(error);
    }
  }

  /** Allocate the next correlation id. Useful when the caller must embed the id
   *  in a transferable-bearing message it builds itself. */
  nextId(): number {
    return this._nextId++;
  }

  /**
   * Send a correlated request and resolve with its matching response. `build`
   * receives the freshly allocated id so it can embed it in the message.
   *
   * `opts.timeoutMs` / `opts.signal` are opt-in escape hatches for a worker
   * that never replies; with neither (and no bridge-level default) the request
   * waits forever, as it always has.
   */
  request(
    build: (id: number) => unknown,
    transfer?: Transferable[],
    opts?: WorkerRequestOptions<TRes>,
  ): Promise<TRes> {
    const id = this._nextId++;
    const timeoutMs = opts?.timeoutMs === false ? undefined : (opts?.timeoutMs ?? this._opts.timeoutMs);
    const signal = opts?.signal;
    const notifyCancel = (reason: 'timeout' | 'abort'): void => {
      try {
        opts?.onCancel?.(id, reason);
      } catch {
        // Cancellation has already settled locally; notification is best-effort.
      }
    };
    return new Promise<TRes>((resolve, reject) => {
      if (this._terminated) {
        reject(new Error('Worker terminated'));
        return;
      }
      if (this._failure) {
        reject(this._failure);
        return;
      }
      // Already-aborted signal: reject synchronously, register nothing.
      if (signal?.aborted) {
        notifyCancel('abort');
        reject(abortError());
        return;
      }

      let timer: ReturnType<typeof setTimeout> | undefined;
      let onAbort: (() => void) | undefined;

      // Runs once, on the first settling path (response, timeout, abort,
      // worker-error, or terminate), to release the timer and abort listener.
      const cleanup = (): void => {
        if (timer !== undefined) {
          clearTimeout(timer);
          timer = undefined;
        }
        if (onAbort && signal) {
          signal.removeEventListener('abort', onAbort);
          onAbort = undefined;
        }
      };

      this._pending.set(id, { resolve, reject, cleanup, onOrphanedResponse: opts?.onOrphanedResponse });

      if (timeoutMs !== undefined) {
        timer = setTimeout(() => {
          const cb = this._pending.get(id);
          if (!cb) return;
          this._pending.delete(id);
          if (cb.onOrphanedResponse) this._orphaned.set(id, cb.onOrphanedResponse);
          cb.cleanup();
          notifyCancel('timeout');
          cb.reject(new Error(`worker request timed out after ${timeoutMs}ms`));
        }, timeoutMs);
      }

      if (signal) {
        onAbort = (): void => {
          const cb = this._pending.get(id);
          if (!cb) return;
          this._pending.delete(id);
          if (cb.onOrphanedResponse) this._orphaned.set(id, cb.onOrphanedResponse);
          cb.cleanup();
          notifyCancel('abort');
          cb.reject(abortError());
        };
        signal.addEventListener('abort', onAbort);
      }

      try {
        this._worker.postMessage(build(id), transfer);
      } catch (error) {
        const cb = this._pending.get(id);
        if (!cb) return;
        this._pending.delete(id);
        cb.cleanup();
        cb.reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  }

  /**
   * Return a typed protocol view backed by this bridge. This installs no new
   * listeners and allocates request ids from the same correlation namespace.
   * The protocol remains responsible for validating its response envelope.
   */
  transport<TProtocolResponse extends TRes>(
    isResponse: (response: TRes) => response is TProtocolResponse,
  ): WorkerBridgeTransport<TProtocolResponse> {
    return {
      request: (build, transfer, options) =>
        this.request(build, transfer, {
          ...options,
          onOrphanedResponse: options?.onOrphanedResponse
            ? (response) => {
                if (isResponse(response)) options.onOrphanedResponse?.(response);
              }
            : undefined,
        }).then((response) => {
          if (!isResponse(response)) {
            throw new Error('worker response does not match the selected transport');
          }
          return response;
        }),
      forgetOrphaned: (ids) => this.forgetOrphaned(ids),
      terminate: () => this.terminate(),
    };
  }

  /** Fire-and-forget message with no correlation (e.g. the `init` message). */
  post(message: unknown, transfer?: Transferable[]): void {
    this._worker.postMessage(message, transfer);
  }

  /**
   * Forget canceled-request tombstones after the worker has accepted the
   * ordered lifecycle barrier for their owner.
   */
  forgetOrphaned(ids: Iterable<number>): void {
    for (const id of ids) this._orphaned.delete(id);
  }

  get orphanedRequestCount(): number {
    return this._orphaned.size;
  }

  /** Terminate once and reject every still-pending request. Repeated calls are no-ops. */
  terminate(): void {
    if (this._terminated) return;
    this._terminated = true;
    this._worker.removeEventListener('message', this._handle);
    this._worker.removeEventListener('messageerror', this._handleWorkerError);
    this._worker.removeEventListener('error', this._handleWorkerError);
    this._worker.terminate();
    this._rejectAll(new Error('Worker terminated'));
    this._orphaned.clear();
  }
}
