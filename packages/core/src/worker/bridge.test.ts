import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { WorkerBridge, type WorkerLike } from './bridge.js';

/** Minimal in-memory Worker stand-in. Records posted messages and lets a test
 *  deliver responses synchronously via {@link respond}. */
class FakeWorker implements WorkerLike {
  posted: unknown[] = [];
  transfers: (Transferable[] | undefined)[] = [];
  terminated = false;
  private messageListeners = new Set<(e: MessageEvent) => void>();
  private errorListeners = new Set<(e: ErrorEvent) => void>();
  private messageErrorListeners = new Set<(e: MessageEvent) => void>();

  postMessage(message: unknown, transfer?: Transferable[]): void {
    this.posted.push(message);
    this.transfers.push(transfer);
  }
  addEventListener(
    type: 'message' | 'messageerror' | 'error',
    listener: (e: never) => void,
  ): void {
    if (type === 'message') this.messageListeners.add(listener as (e: MessageEvent) => void);
    else if (type === 'error') this.errorListeners.add(listener as (e: ErrorEvent) => void);
    else this.messageErrorListeners.add(listener as (e: MessageEvent) => void);
  }
  removeEventListener(
    type: 'message' | 'messageerror' | 'error',
    listener: (e: never) => void,
  ): void {
    if (type === 'message') this.messageListeners.delete(listener as (e: MessageEvent) => void);
    else if (type === 'error') this.errorListeners.delete(listener as (e: ErrorEvent) => void);
    else this.messageErrorListeners.delete(listener as (e: MessageEvent) => void);
  }
  terminate(): void {
    this.terminated = true;
  }
  /** Deliver a message to all listeners, like the worker posting back. */
  respond(data: unknown): void {
    for (const l of this.messageListeners) l({ data } as MessageEvent);
  }
  /** Fire the worker's `error` event (uncaught exception / load failure). */
  emitError(message?: string): void {
    for (const l of this.errorListeners) l({ type: 'error', message } as ErrorEvent);
  }
  emitMessageError(): void {
    for (const l of this.messageErrorListeners) l({ type: 'messageerror' } as MessageEvent);
  }
  /** Number of live `message` listeners — asserts the bridge cleans up. */
  get messageListenerCount(): number {
    return this.messageListeners.size;
  }
}

interface Res {
  id?: number;
  kind: 'ready' | 'ok' | 'error';
  value?: string;
  message?: string;
}

function makeBridge(worker: FakeWorker, onUnsolicited?: (r: Res) => void) {
  return new WorkerBridge<Res>(worker, {
    correlate: (r) => r.id,
    toError: (r) => (r.kind === 'error' ? (r.message ?? 'error') : undefined),
    onUnsolicited,
  });
}

describe('WorkerBridge', () => {
  it('shares one listener and id namespace across typed protocol transports', async () => {
    type Mixed =
      | { type: 'ok'; id: number; value: string }
      | { kind: 'chunk'; requestId: number; value: string };
    const w = new FakeWorker();
    const bridge = new WorkerBridge<Mixed>(w, {
      correlate: (response) => 'requestId' in response ? response.requestId : response.id,
    });
    const pull = bridge.transport(
      (response): response is Extract<Mixed, { kind: 'chunk' }> => 'requestId' in response,
    );
    const ordinaryRequest = bridge.request((id) => ({ type: 'ordinary', id }));
    const pullRequest = pull.request((requestId) => ({ kind: 'pull', requestId }));
    const ordinaryId = (w.posted[0] as { id: number }).id;
    const requestId = (w.posted[1] as { requestId: number }).requestId;

    expect(w.messageListenerCount).toBe(1);
    expect(requestId).not.toBe(ordinaryId);
    w.respond({ kind: 'chunk', requestId, value: 'rows' });
    w.respond({ type: 'ok', id: ordinaryId, value: 'metadata' });
    await expect(pullRequest).resolves.toMatchObject({ value: 'rows' });
    await expect(ordinaryRequest).resolves.toMatchObject({ value: 'metadata' });
  });

  it('rejects a response outside a typed transport at runtime', async () => {
    type Mixed = { type: 'ok'; id: number } | { kind: 'chunk'; requestId: number };
    const w = new FakeWorker();
    const bridge = new WorkerBridge<Mixed>(w, {
      correlate: (response) => 'requestId' in response ? response.requestId : response.id,
    });
    const pull = bridge.transport(
      (response): response is Extract<Mixed, { kind: 'chunk' }> => 'requestId' in response,
    );
    const pending = pull.request((requestId) => ({ kind: 'pull', requestId }));
    w.respond({ type: 'ok', id: (w.posted[0] as { requestId: number }).requestId });
    await expect(pending).rejects.toThrow('does not match the selected transport');
  });

  it('correlates a response to its request by id', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const sentId = (w.posted[0] as { id: number }).id;
    w.respond({ id: sentId, kind: 'ok', value: 'A' });
    await expect(p).resolves.toMatchObject({ value: 'A' });
  });

  it('routes concurrent responses to the matching request, even out of order', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p1 = bridge.request((id) => ({ kind: 'parse', id }));
    const p2 = bridge.request((id) => ({ kind: 'parse', id }));
    const id1 = (w.posted[0] as { id: number }).id;
    const id2 = (w.posted[1] as { id: number }).id;
    expect(id1).not.toBe(id2);
    // Respond to the second request first.
    w.respond({ id: id2, kind: 'ok', value: 'second' });
    w.respond({ id: id1, kind: 'ok', value: 'first' });
    await expect(p1).resolves.toMatchObject({ value: 'first' });
    await expect(p2).resolves.toMatchObject({ value: 'second' });
  });

  it('rejects only the matching request on an error response', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p1 = bridge.request((id) => ({ kind: 'parse', id }));
    const p2 = bridge.request((id) => ({ kind: 'parse', id }));
    const id1 = (w.posted[0] as { id: number }).id;
    const id2 = (w.posted[1] as { id: number }).id;
    w.respond({ id: id1, kind: 'error', message: 'boom' });
    w.respond({ id: id2, kind: 'ok', value: 'fine' });
    await expect(p1).rejects.toThrow('boom');
    await expect(p2).resolves.toMatchObject({ value: 'fine' });
  });

  it('preserves a reconstructed Error returned by the response adapter', async () => {
    const w = new FakeWorker();
    const bridge = new WorkerBridge<Res>(w, {
      correlate: (r) => r.id,
      toError: (r) => {
        if (r.kind !== 'error') return undefined;
        return Object.assign(new Error(r.message), {
          name: 'StructuredWorkerError',
          code: 'WORKER_CODE',
        });
      },
    });
    const pending = bridge.request((id) => ({ kind: 'parse', id }));
    const id = (w.posted[0] as { id: number }).id;

    w.respond({ id, kind: 'error', message: 'boom' });

    await expect(pending).rejects.toMatchObject({
      name: 'StructuredWorkerError',
      message: 'boom',
      code: 'WORKER_CODE',
    });
  });

  it('does not resolve or hang the wrong request when an unknown id arrives', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const id = (w.posted[0] as { id: number }).id;
    w.respond({ id: 9999, kind: 'ok', value: 'stray' }); // unknown id: ignored
    w.respond({ id, kind: 'ok', value: 'real' });
    await expect(p).resolves.toMatchObject({ value: 'real' });
  });

  it('forwards unsolicited messages (no id) to onUnsolicited instead of a pending request', async () => {
    const w = new FakeWorker();
    const seen: Res[] = [];
    const bridge = makeBridge(w, (r) => seen.push(r));
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    w.respond({ kind: 'ready' }); // no id
    const id = (w.posted[0] as { id: number }).id;
    w.respond({ id, kind: 'ok', value: 'done' });
    await expect(p).resolves.toMatchObject({ value: 'done' });
    expect(seen).toEqual([{ kind: 'ready' }]);
  });

  it('rejects all pending requests when terminated', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    bridge.terminate();
    expect(w.terminated).toBe(true);
    await expect(p).rejects.toThrow(/terminated/i);
  });

  it('passes the transfer list through to postMessage', () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const buf = new ArrayBuffer(8);
    bridge.request((id) => ({ kind: 'parse', id, buffer: buf }), [buf]);
    expect(w.transfers[0]).toEqual([buf]);
  });

  it('rejects and removes pending state when postMessage throws synchronously', async () => {
    vi.useFakeTimers();
    try {
      const w = new FakeWorker();
      w.postMessage = vi.fn(() => {
        throw new DOMException('could not clone', 'DataCloneError');
      });
      const bridge = makeBridge(w);
      const pending = bridge.request((id) => ({ kind: 'parse', id }), undefined, { timeoutMs: 10 });
      await expect(pending).rejects.toMatchObject({ name: 'DataCloneError' });
      expect(vi.getTimerCount()).toBe(0);
      bridge.terminate();
      expect(w.terminated).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });

  it('post() sends a fire-and-forget message without allocating an id', () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    bridge.post({ kind: 'init', wasmUrl: 'x' });
    expect(w.posted[0]).toEqual({ kind: 'init', wasmUrl: 'x' });
    // The next request should still start from id 1.
    bridge.request((id) => ({ kind: 'parse', id }));
    expect((w.posted[1] as { id: number }).id).toBe(1);
  });

  it('ignores a duplicate/late response after the request already settled', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const id = (w.posted[0] as { id: number }).id;
    w.respond({ id, kind: 'ok', value: 'first' });
    await expect(p).resolves.toMatchObject({ value: 'first' });
    // A stray duplicate must not throw or affect anything.
    expect(() => w.respond({ id, kind: 'ok', value: 'dup' })).not.toThrow();
  });

  it('disposes an unknown or late correlated response without affecting requests', async () => {
    const w = new FakeWorker();
    const dispose = vi.fn();
    const bridge = new WorkerBridge<Res>(w, {
      correlate: (r) => r.id,
      onOrphanedResponse: dispose,
    });
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const id = (w.posted[0] as { id: number }).id;

    w.respond({ id: 999, kind: 'ok', value: 'unknown' });
    w.respond({ id, kind: 'ok', value: 'real' });
    await p;
    w.respond({ id, kind: 'ok', value: 'late' });

    expect(dispose).toHaveBeenCalledTimes(2);
    expect(dispose).toHaveBeenNthCalledWith(1, { id: 999, kind: 'ok', value: 'unknown' });
    expect(dispose).toHaveBeenNthCalledWith(2, { id, kind: 'ok', value: 'late' });
  });

  describe('timeout', () => {
    beforeEach(() => {
      vi.useFakeTimers();
    });
    afterEach(() => {
      vi.restoreAllMocks();
      vi.useRealTimers();
    });

    it('rejects a request whose worker never answers within timeoutMs', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        timeoutMs: 1000,
      });
      const rejects = expect(p).rejects.toThrow(/timed out after 1000ms/);
      await vi.advanceTimersByTimeAsync(1000);
      await rejects;
    });

    it('notifies the caller when timeout cancels a correlated request', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const onCancel = vi.fn();
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        timeoutMs: 10,
        onCancel,
      });
      const id = (w.posted[0] as { id: number }).id;
      const rejects = expect(p).rejects.toThrow(/timed out/);
      await vi.advanceTimersByTimeAsync(10);
      await rejects;
      expect(onCancel).toHaveBeenCalledOnce();
      expect(onCancel).toHaveBeenCalledWith(id, 'timeout');
    });

    it('honours a bridge-level default timeoutMs', async () => {
      const w = new FakeWorker();
      const bridge = new WorkerBridge<Res>(w, {
        correlate: (r) => r.id,
        toError: (r) => (r.kind === 'error' ? (r.message ?? 'error') : undefined),
        timeoutMs: 500,
      });
      const p = bridge.request((id) => ({ kind: 'parse', id }));
      const rejects = expect(p).rejects.toThrow(/timed out after 500ms/);
      await vi.advanceTimersByTimeAsync(500);
      await rejects;
    });

    it('lets a per-request timeoutMs override the bridge default', async () => {
      const w = new FakeWorker();
      const bridge = new WorkerBridge<Res>(w, {
        correlate: (r) => r.id,
        timeoutMs: 5000,
      });
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        timeoutMs: 100,
      });
      const rejects = expect(p).rejects.toThrow(/timed out after 100ms/);
      await vi.advanceTimersByTimeAsync(100);
      await rejects;
    });

    it('clears the timer when the response arrives first (no late reject/leak)', async () => {
      const clearSpy = vi.spyOn(globalThis, 'clearTimeout');
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        timeoutMs: 1000,
      });
      const id = (w.posted[0] as { id: number }).id;
      w.respond({ id, kind: 'ok', value: 'quick' });
      await expect(p).resolves.toMatchObject({ value: 'quick' });
      expect(clearSpy).toHaveBeenCalled();
      // No pending timer should fire after settle.
      expect(vi.getTimerCount()).toBe(0);
    });

    it('creates no timer at all when neither per-request nor bridge default is set', () => {
      const setSpy = vi.spyOn(globalThis, 'setTimeout');
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      bridge.request((id) => ({ kind: 'parse', id }));
      expect(setSpy).not.toHaveBeenCalled();
      expect(vi.getTimerCount()).toBe(0);
    });
  });

  describe('AbortSignal', () => {
    it('rejects immediately with an AbortError when the signal is already aborted', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        signal: AbortSignal.abort(),
      });
      await expect(p).rejects.toMatchObject({ name: 'AbortError' });
      // Nothing was registered, so no request is left pending and no message posted.
      expect(w.posted).toHaveLength(0);
    });

    it('notifies cancellation for an already-aborted request without posting it', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const onCancel = vi.fn();
      await expect(
        bridge.request((id) => ({ kind: 'parse', id }), undefined, {
          signal: AbortSignal.abort(),
          onCancel,
        }),
      ).rejects.toMatchObject({ name: 'AbortError' });
      expect(onCancel).toHaveBeenCalledWith(1, 'abort');
      expect(w.posted).toHaveLength(0);
    });

    it('rejects and detaches the listener when the signal aborts mid-flight', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const ctrl = new AbortController();
      const removeSpy = vi.spyOn(ctrl.signal, 'removeEventListener');
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        signal: ctrl.signal,
      });
      ctrl.abort();
      await expect(p).rejects.toMatchObject({ name: 'AbortError' });
      expect(removeSpy).toHaveBeenCalledWith('abort', expect.any(Function));
    });

    it('notifies the caller when abort cancels a correlated request', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const ctrl = new AbortController();
      const onCancel = vi.fn();
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        signal: ctrl.signal,
        onCancel,
      });
      const id = (w.posted[0] as { id: number }).id;
      ctrl.abort();
      await expect(p).rejects.toMatchObject({ name: 'AbortError' });
      expect(onCancel).toHaveBeenCalledWith(id, 'abort');
    });

    it('detaches the abort listener when the response arrives before any abort', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const ctrl = new AbortController();
      const removeSpy = vi.spyOn(ctrl.signal, 'removeEventListener');
      const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
        signal: ctrl.signal,
      });
      const id = (w.posted[0] as { id: number }).id;
      w.respond({ id, kind: 'ok', value: 'done' });
      await expect(p).resolves.toMatchObject({ value: 'done' });
      expect(removeSpy).toHaveBeenCalledWith('abort', expect.any(Function));
      // A later abort must not throw or touch the already-settled request.
      expect(() => ctrl.abort()).not.toThrow();
    });
  });

  describe('worker error events', () => {
    it("rejects all pending requests on the worker's error event", async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const p1 = bridge.request((id) => ({ kind: 'parse', id }));
      const p2 = bridge.request((id) => ({ kind: 'parse', id }));
      w.emitError('boom in worker');
      await expect(p1).rejects.toThrow(/Worker error.*boom in worker/);
      await expect(p2).rejects.toThrow(/Worker error.*boom in worker/);
      // The message listener should be torn down for each settled request.
      expect(w.messageListenerCount).toBe(1); // only the bridge's own handler remains
    });

    it('rejects future cleanup requests immediately after the worker becomes unusable', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      w.emitError('worker is gone');
      const posted = w.posted.length;
      await expect(bridge.request((id) => ({ kind: 'cancel', id }))).rejects.toThrow(
        /Worker error.*worker is gone/,
      );
      expect(w.posted).toHaveLength(posted);
    });

    it('rejects pending requests on a messageerror (undeserializable response) too', async () => {
      const w = new FakeWorker();
      const bridge = makeBridge(w);
      const p = bridge.request((id) => ({ kind: 'parse', id }));
      // messageerror events carry no `.message`; the reject is still generic.
      w.emitMessageError();
      await expect(p).rejects.toThrow(/Worker error/);
    });

    it('explains an empty worker load error from an opaque origin', async () => {
      vi.stubGlobal('origin', 'null');
      try {
        const w = new FakeWorker();
        const bridge = makeBridge(w);
        const pending = bridge.request((id) => ({ kind: 'parse', id }));
        w.emitError('');
        await expect(pending).rejects.toThrow(/sandboxed iframe.*allow-same-origin/);

        const invalidResponseWorker = new FakeWorker();
        const invalidResponse = makeBridge(invalidResponseWorker)
          .request((id) => ({ kind: 'parse', id }));
        invalidResponseWorker.emitMessageError();
        await expect(invalidResponse).rejects.toThrow(/^Worker error$/);
      } finally {
        vi.unstubAllGlobals();
      }
    });

    it('clears a pending timeout when the worker errors (no leaked timer)', async () => {
      vi.useFakeTimers();
      try {
        const clearSpy = vi.spyOn(globalThis, 'clearTimeout');
        const w = new FakeWorker();
        const bridge = makeBridge(w);
        const p = bridge.request((id) => ({ kind: 'parse', id }), undefined, {
          timeoutMs: 1000,
        });
        w.emitError('crash');
        await expect(p).rejects.toThrow(/Worker error/);
        expect(clearSpy).toHaveBeenCalled();
        expect(vi.getTimerCount()).toBe(0);
      } finally {
        vi.restoreAllMocks();
        vi.useRealTimers();
      }
    });
  });
});
