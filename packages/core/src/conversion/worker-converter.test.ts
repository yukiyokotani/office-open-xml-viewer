import { describe, expect, it, vi } from 'vitest';
import { Buffer } from 'node:buffer';
import { LegacyOfficeConversionError, type LegacyOfficeConverter } from './legacy-office.js';
import {
  createDisposableWorkerLegacyOfficeConverter,
  installLegacyOfficeConversionWorkerHandler,
  type LegacyOfficeConversionWorker,
  type LegacyOfficeConversionWorkerScope,
  type LegacyOfficeWorkerRequest,
  type LegacyOfficeWorkerResponse,
} from './worker-converter.js';

type Listener = (event: MessageEvent<LegacyOfficeWorkerResponse>) => void;

class FakeWorker implements LegacyOfficeConversionWorker {
  readonly messages: Array<Readonly<{ request: LegacyOfficeWorkerRequest; transfer: Transferable[] }>> = [];
  readonly listeners = new Set<Listener>();
  readonly errors = new Set<(event: ErrorEvent) => void>();
  readonly messageErrors = new Set<(event: MessageEvent) => void>();
  terminate = vi.fn();

  postMessage(request: LegacyOfficeWorkerRequest, transfer: Transferable[]): void {
    this.messages.push({ request, transfer });
  }

  addEventListener(type: string, listener: EventListener): void {
    if (type === 'message') this.listeners.add(listener as unknown as Listener);
    if (type === 'error') this.errors.add(listener as unknown as (event: ErrorEvent) => void);
    if (type === 'messageerror') this.messageErrors.add(listener as unknown as (event: MessageEvent) => void);
  }

  removeEventListener(type: string, listener: EventListener): void {
    if (type === 'message') this.listeners.delete(listener as unknown as Listener);
    if (type === 'error') this.errors.delete(listener as unknown as (event: ErrorEvent) => void);
    if (type === 'messageerror') this.messageErrors.delete(listener as unknown as (event: MessageEvent) => void);
  }

  emit(response: LegacyOfficeWorkerResponse): void {
    for (const listener of this.listeners) listener(new MessageEvent('message', { data: response }));
  }
}

describe('createDisposableWorkerLegacyOfficeConverter', () => {
  it('transfers one exact input buffer and terminates the worker after success', async () => {
    const worker = new FakeWorker();
    const converter = createDisposableWorkerLegacyOfficeConverter(() => worker);
    const input = new Uint8Array([1, 2, 3]);

    const pending = converter.convert({
      bytes: input,
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });
    expect(worker.messages).toHaveLength(1);
    expect(worker.messages[0]?.request).toMatchObject({
      type: 'convert',
      requestId: 1,
      from: 'doc',
      to: 'docx',
      bytes: input.buffer,
    });
    expect(worker.messages[0]?.transfer).toEqual([input.buffer]);

    const output = new Uint8Array([4, 5, 6]);
    worker.emit({
      type: 'converted',
      requestId: 1,
      bytes: output.buffer,
      engine: 'legacy-wasm',
      engineVersion: '0.1.0',
      outputSha256: 'b'.repeat(64),
      warnings: ['shape omitted'],
    });

    await expect(pending).resolves.toEqual({
      bytes: output.buffer,
      engine: 'legacy-wasm',
      engineVersion: '0.1.0',
      outputSha256: 'b'.repeat(64),
      warnings: ['shape omitted'],
    });
    expect(worker.terminate).toHaveBeenCalledOnce();
    expect(worker.listeners).toHaveLength(0);
    expect(worker.errors).toHaveLength(0);
    expect(worker.messageErrors).toHaveLength(0);
  });

  it('runs one disposable worker at a time by default', async () => {
    const workers: FakeWorker[] = [];
    const converter = createDisposableWorkerLegacyOfficeConverter(() => {
      const worker = new FakeWorker();
      workers.push(worker);
      return worker;
    });
    const request = () => converter.convert({
      bytes: new Uint8Array([1]),
      from: 'doc' as const,
      to: 'docx' as const,
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });

    const first = request();
    const second = request();
    expect(workers).toHaveLength(1);

    workers[0]?.emit({ type: 'converted', requestId: 1, bytes: new ArrayBuffer(0) });
    await expect(first).resolves.toMatchObject({ bytes: expect.any(ArrayBuffer) });
    await vi.waitFor(() => expect(workers).toHaveLength(2));
    workers[1]?.emit({ type: 'converted', requestId: 1, bytes: new ArrayBuffer(0) });
    await expect(second).resolves.toMatchObject({ bytes: expect.any(ArrayBuffer) });
  });

  it('rejects beyond a configured bounded queue', async () => {
    const worker = new FakeWorker();
    const converter = createDisposableWorkerLegacyOfficeConverter(
      () => worker,
      { maxConcurrency: 1, maxQueuedConversions: 0 },
    );
    const request = () => converter.convert({
      bytes: new Uint8Array([1]),
      from: 'doc' as const,
      to: 'docx' as const,
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });

    const active = request();
    await expect(request()).rejects.toMatchObject({ reason: 'capacity-exceeded' });
    worker.emit({ type: 'converted', requestId: 1, bytes: new ArrayBuffer(0) });
    await expect(active).resolves.toMatchObject({ bytes: expect.any(ArrayBuffer) });
  });

  it('removes a cancelled request while it is waiting in the queue', async () => {
    const workers: FakeWorker[] = [];
    const converter = createDisposableWorkerLegacyOfficeConverter(() => {
      const worker = new FakeWorker();
      workers.push(worker);
      return worker;
    });
    const active = converter.convert({
      bytes: new Uint8Array([1]),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });
    const queuedController = new AbortController();
    const queued = converter.convert({
      bytes: new Uint8Array([2]),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
      signal: queuedController.signal,
    });

    queuedController.abort();
    await expect(queued).rejects.toMatchObject({ reason: 'aborted' });
    workers[0]?.emit({ type: 'converted', requestId: 1, bytes: new ArrayBuffer(0) });
    await active;
    expect(workers).toHaveLength(1);
  });

  it('terminates immediately and returns a typed abort when its signal fires', async () => {
    const worker = new FakeWorker();
    const controller = new AbortController();
    const converter = createDisposableWorkerLegacyOfficeConverter(() => worker);
    const pending = converter.convert({
      bytes: new Uint8Array([1]),
      from: 'xls',
      to: 'xlsx',
      maxOutputBytes: 1024,
      signal: controller.signal,
    });

    controller.abort();

    await expect(pending).rejects.toMatchObject({
      code: 'legacy-office-conversion',
      reason: 'aborted',
      from: 'xls',
      to: 'xlsx',
    });
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it('transfers only the addressed bytes of a Node Buffer subview', async () => {
    const worker = new FakeWorker();
    const converter = createDisposableWorkerLegacyOfficeConverter(() => worker);
    const pooled = Buffer.from([9, 1, 2, 3, 9]);

    const pending = converter.convert({
      bytes: pooled.subarray(1, 4),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });

    const transferred = worker.messages[0]?.request.bytes;
    expect(transferred?.byteLength).toBe(3);
    expect(new Uint8Array(transferred as ArrayBuffer)).toEqual(new Uint8Array([1, 2, 3]));
    worker.emit({ type: 'converted', requestId: 1, bytes: new ArrayBuffer(0) });
    await expect(pending).resolves.toMatchObject({ bytes: expect.any(ArrayBuffer) });
  });

  it('preserves a content-free unsupported-input response', async () => {
    const worker = new FakeWorker();
    const converter = createDisposableWorkerLegacyOfficeConverter(() => worker);
    const pending = converter.convert({
      bytes: new Uint8Array([1]),
      from: 'ppt',
      to: 'pptx',
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });

    worker.emit({ type: 'conversion-error', requestId: 1, reason: 'unsupported-input' });

    await expect(pending).rejects.toMatchObject({ reason: 'unsupported-input' });
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it('preserves an output budget rejection from the converter worker', async () => {
    const worker = new FakeWorker();
    const converter = createDisposableWorkerLegacyOfficeConverter(() => worker);
    const pending = converter.convert({
      bytes: new Uint8Array([1]),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 2,
      signal: new AbortController().signal,
    });

    worker.emit({ type: 'conversion-error', requestId: 1, reason: 'output-too-large' });

    await expect(pending).rejects.toMatchObject({ reason: 'output-too-large' });
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it('maps worker crashes and malformed responses to sanitized typed failures', async () => {
    const crashWorker = new FakeWorker();
    const crashConverter = createDisposableWorkerLegacyOfficeConverter(() => crashWorker);
    const crashed = crashConverter.convert({
      bytes: new Uint8Array([1]),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });
    for (const listener of crashWorker.errors) {
      listener({
        message: 'private filename.doc',
        preventDefault: vi.fn(),
      } as unknown as ErrorEvent);
    }
    await expect(crashed).rejects.toSatisfy((error: unknown) => {
      expect(error).toBeInstanceOf(LegacyOfficeConversionError);
      expect(error).toMatchObject({ reason: 'failed' });
      expect(String(error)).not.toContain('private filename.doc');
      return true;
    });

    const malformedWorker = new FakeWorker();
    const malformedConverter = createDisposableWorkerLegacyOfficeConverter(() => malformedWorker);
    const malformed = malformedConverter.convert({
      bytes: new Uint8Array([1]),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
      signal: new AbortController().signal,
    });
    malformedWorker.emit({ type: 'converted', requestId: 1, bytes: {} as ArrayBuffer });
    await expect(malformed).rejects.toMatchObject({ reason: 'invalid-output' });
    expect(malformedWorker.terminate).toHaveBeenCalledOnce();
  });
});

describe('installLegacyOfficeConversionWorkerHandler', () => {
  it('runs a same-family converter once and transfers an exact output buffer', async () => {
    let listener: EventListener | undefined;
    const posted: Array<{ message: LegacyOfficeWorkerResponse; transfer: Transferable[] }> = [];
    const scope: LegacyOfficeConversionWorkerScope = {
      postMessage: (message, transfer) => posted.push({ message, transfer }),
      addEventListener: (_type, next) => { listener = next; },
      removeEventListener: (_type, next) => {
        if (listener === next) listener = undefined;
      },
    };
    const backing = Buffer.from([9, 4, 5, 6, 9]);
    const convert = vi.fn<LegacyOfficeConverter['convert']>(async () => ({
      bytes: backing.subarray(1, 4),
      engine: 'wasm-test',
    }));
    const cleanup = installLegacyOfficeConversionWorkerHandler(scope, { convert });

    listener?.(new MessageEvent('message', { data: {
      type: 'convert',
      requestId: 1,
      bytes: new Uint8Array([1, 2]).buffer,
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
    } satisfies LegacyOfficeWorkerRequest }));
    await vi.waitFor(() => expect(posted).toHaveLength(1));

    const response = posted[0]?.message;
    expect(response).toMatchObject({ type: 'converted', engine: 'wasm-test' });
    if (response?.type !== 'converted') throw new Error('expected conversion');
    expect(new Uint8Array(response.bytes)).toEqual(new Uint8Array([4, 5, 6]));
    expect(posted[0]?.transfer).toEqual([response.bytes]);
    expect(convert).toHaveBeenCalledOnce();

    // A disposable worker accepts no second conversion before host termination.
    listener?.(new MessageEvent('message', { data: {
      type: 'convert',
      requestId: 1,
      bytes: new ArrayBuffer(0),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024,
    } satisfies LegacyOfficeWorkerRequest }));
    await Promise.resolve();
    expect(convert).toHaveBeenCalledOnce();
    cleanup();
    expect(listener).toBeUndefined();
  });

  it('rejects cross-family protocol messages before invoking converter code', async () => {
    let listener: EventListener | undefined;
    const postMessage = vi.fn();
    const convert = vi.fn<LegacyOfficeConverter['convert']>();
    installLegacyOfficeConversionWorkerHandler({
      postMessage,
      addEventListener: (_type, next) => { listener = next; },
      removeEventListener: vi.fn(),
    }, { convert });

    listener?.(new MessageEvent('message', { data: {
      type: 'convert',
      requestId: 1,
      bytes: new ArrayBuffer(0),
      from: 'doc',
      to: 'xlsx',
      maxOutputBytes: 1024,
    } }));
    await Promise.resolve();

    expect(convert).not.toHaveBeenCalled();
    expect(postMessage).toHaveBeenCalledWith(
      { type: 'conversion-error', requestId: 1, reason: 'failed' },
      [],
    );
  });

  it('rejects oversized worker output before copying or transferring it', async () => {
    let listener: EventListener | undefined;
    const postMessage = vi.fn();
    installLegacyOfficeConversionWorkerHandler({
      postMessage,
      addEventListener: (_type, next) => { listener = next; },
      removeEventListener: vi.fn(),
    }, {
      convert: async () => ({ bytes: new Uint8Array(3) }),
    });

    listener?.(new MessageEvent('message', { data: {
      type: 'convert',
      requestId: 1,
      bytes: new ArrayBuffer(0),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 2,
    } satisfies LegacyOfficeWorkerRequest }));
    await vi.waitFor(() => expect(postMessage).toHaveBeenCalledOnce());

    expect(postMessage).toHaveBeenCalledWith(
      { type: 'conversion-error', requestId: 1, reason: 'output-too-large' },
      [],
    );
  });
});
