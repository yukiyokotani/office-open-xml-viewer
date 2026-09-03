import {
  HARD_MAX_LEGACY_CONVERSION_BYTES,
  type LegacyOfficeConversionInput,
  type LegacyOfficeConversionResult,
  type LegacyOfficeConverter,
} from './legacy-office.js';
import {
  LegacyOfficeConversionError,
  type LegacyOfficeConversionFailureReason,
  type LegacyOfficeFormat,
} from './legacy-office-error.js';
import type { OoxmlFormat } from '../errors/ooxml-error.js';

/** One-request protocol sent to a disposable converter Worker. */
export interface LegacyOfficeWorkerRequest {
  readonly type: 'convert';
  readonly requestId: 1;
  readonly bytes: ArrayBuffer;
  readonly from: LegacyOfficeFormat;
  readonly to: OoxmlFormat;
  readonly maxOutputBytes: number;
}

export type LegacyOfficeWorkerResponse =
  | {
      readonly type: 'converted';
      readonly requestId: 1;
      readonly bytes: ArrayBuffer;
      readonly engine?: string;
      readonly engineVersion?: string;
      readonly outputSha256?: string;
      readonly warnings?: readonly string[];
    }
  | {
      readonly type: 'conversion-error';
      readonly requestId: 1;
      /** Worker responses intentionally carry no free-form error message. */
      readonly reason: 'unsupported-input' | 'output-too-large' | 'failed';
    };

/** Minimal browser Worker surface accepted by the disposable adapter. */
export interface LegacyOfficeConversionWorker {
  postMessage(message: LegacyOfficeWorkerRequest, transfer: Transferable[]): void;
  addEventListener(type: string, listener: EventListener): void;
  removeEventListener(type: string, listener: EventListener): void;
  terminate(): void;
}

export type LegacyOfficeConversionWorkerFactory = () => LegacyOfficeConversionWorker;

export interface LegacyOfficeConversionWorkerAdapterOptions {
  /** Simultaneously live disposable Workers. Default: 1; hard maximum: 16. */
  readonly maxConcurrency?: number;
  /** Waiting requests retained by this adapter. Default: 4; hard maximum: 64. */
  readonly maxQueuedConversions?: number;
}

/** Minimal dedicated-worker global surface used by the protocol host. */
export interface LegacyOfficeConversionWorkerScope {
  postMessage(message: LegacyOfficeWorkerResponse, transfer: Transferable[]): void;
  addEventListener(type: 'message', listener: EventListener): void;
  removeEventListener(type: 'message', listener: EventListener): void;
}

/**
 * Adapt a lazily-created browser Worker to {@link LegacyOfficeConverter}.
 *
 * A fresh Worker is created for every conversion. Input ownership is
 * transferred into it, output ownership must be transferred back by the Worker,
 * and the Worker is terminated on every terminal path so converter WASM memory
 * cannot overlap the parser's retained lifetime.
 */
export function createDisposableWorkerLegacyOfficeConverter(
  createWorker: LegacyOfficeConversionWorkerFactory,
  options: LegacyOfficeConversionWorkerAdapterOptions = {},
): LegacyOfficeConverter {
  const maxConcurrency = boundedInteger(options.maxConcurrency, 1, 1, 16, 'maxConcurrency');
  const maxQueuedConversions = boundedInteger(
    options.maxQueuedConversions,
    4,
    0,
    64,
    'maxQueuedConversions',
  );
  let active = 0;
  const queue: QueuedWorkerConversion[] = [];

  const drain = (): void => {
    while (active < maxConcurrency) {
      const queued = queue.shift();
      if (queued === undefined) return;
      queued.input.signal.removeEventListener('abort', queued.onAbort);
      if (queued.input.signal.aborted) {
        queued.reject(conversionFailure('aborted', queued.input));
        continue;
      }
      active++;
      void convertInDisposableWorker(createWorker, queued.input).then(
        (result) => {
          active--;
          queued.resolve(result);
          drain();
        },
        (error: unknown) => {
          active--;
          queued.reject(error);
          drain();
        },
      );
    }
  };

  return {
    convert: (input) => {
      if (input.signal.aborted) {
        return Promise.reject(conversionFailure('aborted', input));
      }
      if (active >= maxConcurrency && queue.length >= maxQueuedConversions) {
        return Promise.reject(conversionFailure('capacity-exceeded', input));
      }
      return new Promise<LegacyOfficeConversionResult>((resolve, reject) => {
        const queued: QueuedWorkerConversion = {
          input,
          resolve,
          reject,
          onAbort: () => {
            const index = queue.indexOf(queued);
            if (index < 0) return;
            queue.splice(index, 1);
            reject(conversionFailure('aborted', input));
          },
        };
        input.signal.addEventListener('abort', queued.onAbort, { once: true });
        queue.push(queued);
        drain();
      });
    },
  };
}

interface QueuedWorkerConversion {
  readonly input: Readonly<LegacyOfficeConversionInput>;
  readonly resolve: (result: LegacyOfficeConversionResult) => void;
  readonly reject: (error: unknown) => void;
  readonly onAbort: () => void;
}

/**
 * Install the matching one-shot protocol inside a converter Worker. A WASM
 * implementation can initialize lazily in its converter and return standalone
 * output bytes; this host transfers those bytes back and never serializes a
 * free-form failure message.
 *
 * The returned cleanup is mainly useful to worker tests. Production workers are
 * terminated by the disposable main-thread adapter after their first response.
 */
export function installLegacyOfficeConversionWorkerHandler(
  scope: LegacyOfficeConversionWorkerScope,
  converter: LegacyOfficeConverter,
): () => void {
  let started = false;
  const handleMessage = async (event: MessageEvent<unknown>): Promise<void> => {
    if (started) return;
    if (!isWorkerRequest(event.data)) {
      if (isResponseObject(event.data)
        && event.data.type === 'convert'
        && event.data.requestId === 1) {
        started = true;
        scope.postMessage({ type: 'conversion-error', requestId: 1, reason: 'failed' }, []);
      }
      return;
    }
    started = true;
    const request = event.data;
    const from = request.from;
    const to = request.to;
    try {
      const result = await converter.convert({
        bytes: new Uint8Array(request.bytes),
        from,
        to,
        maxOutputBytes: request.maxOutputBytes,
        // Cancellation is enforced by terminating this disposable Worker.
        signal: new AbortController().signal,
      });
      const output = workerOutputBuffer(result.bytes, request.maxOutputBytes, from, to);
      const response: LegacyOfficeWorkerResponse = {
        type: 'converted',
        requestId: 1,
        bytes: output,
        ...(result.engine === undefined ? {} : { engine: result.engine }),
        ...(result.engineVersion === undefined ? {} : { engineVersion: result.engineVersion }),
        ...(result.outputSha256 === undefined ? {} : { outputSha256: result.outputSha256 }),
        ...(result.warnings === undefined ? {} : { warnings: result.warnings }),
      };
      scope.postMessage(response, [output]);
    } catch (error) {
      const reason = error instanceof LegacyOfficeConversionError
        && (error.reason === 'unsupported-input' || error.reason === 'output-too-large')
        ? error.reason
        : 'failed';
      scope.postMessage({ type: 'conversion-error', requestId: 1, reason }, []);
    }
  };
  const listener: EventListener = (event) => {
    if (!(event instanceof MessageEvent)) return;
    void handleMessage(event);
  };
  scope.addEventListener('message', listener);
  return () => scope.removeEventListener('message', listener);
}

function convertInDisposableWorker(
  createWorker: LegacyOfficeConversionWorkerFactory,
  input: Readonly<LegacyOfficeConversionInput>,
): Promise<LegacyOfficeConversionResult> {
  if (input.signal.aborted) {
    return Promise.reject(new LegacyOfficeConversionError('aborted', input.from, input.to));
  }

  let worker: LegacyOfficeConversionWorker;
  try {
    worker = createWorker();
  } catch {
    return Promise.reject(new LegacyOfficeConversionError('failed', input.from, input.to));
  }

  return new Promise<LegacyOfficeConversionResult>((resolve, reject) => {
    let settled = false;
    const finish = (
      action: () => void,
    ): void => {
      if (settled) return;
      settled = true;
      input.signal.removeEventListener('abort', onAbort);
      worker.removeEventListener('message', onMessage);
      worker.removeEventListener('error', onWorkerError);
      worker.removeEventListener('messageerror', onMessageError);
      try {
        worker.terminate();
      } catch {
        // Preserve the conversion outcome if a host-specific terminate throws.
      }
      action();
    };
    const fail = (reason: LegacyOfficeConversionFailureReason): void => {
      finish(() => reject(new LegacyOfficeConversionError(reason, input.from, input.to)));
    };
    const onAbort = (): void => fail('aborted');
    const onMessage: EventListener = (event): void => {
      if (!(event instanceof MessageEvent)) return;
      const response = event.data;
      if (!isResponseObject(response) || response.requestId !== 1) return;
      if (response.type === 'conversion-error') {
        fail(
          response.reason === 'unsupported-input' || response.reason === 'output-too-large'
            ? response.reason
            : 'failed',
        );
        return;
      }
      if (response.type !== 'converted' || !(response.bytes instanceof ArrayBuffer)) {
        fail('invalid-output');
        return;
      }
      const engine = response.engine;
      const engineVersion = response.engineVersion;
      const outputSha256 = response.outputSha256;
      const warnings = response.warnings;
      if (
        (engine !== undefined && typeof engine !== 'string')
        || (engineVersion !== undefined && typeof engineVersion !== 'string')
        || (outputSha256 !== undefined && typeof outputSha256 !== 'string')
        || (warnings !== undefined
          && (!Array.isArray(warnings) || !warnings.every((warning) => typeof warning === 'string')))
      ) {
        fail('invalid-output');
        return;
      }
      const result: LegacyOfficeConversionResult = {
        bytes: response.bytes,
        ...(engine === undefined ? {} : { engine }),
        ...(engineVersion === undefined ? {} : { engineVersion }),
        ...(outputSha256 === undefined ? {} : { outputSha256 }),
        ...(warnings === undefined ? {} : { warnings: warnings as string[] }),
      };
      finish(() => resolve(result));
    };
    const onWorkerError: EventListener = (event): void => {
      event.preventDefault();
      fail('failed');
    };
    const onMessageError: EventListener = () => fail('invalid-output');

    input.signal.addEventListener('abort', onAbort, { once: true });
    worker.addEventListener('message', onMessage);
    worker.addEventListener('error', onWorkerError);
    worker.addEventListener('messageerror', onMessageError);

    try {
      if (input.signal.aborted) {
        onAbort();
        return;
      }
      const bytes = exactTransferableInput(input.bytes);
      worker.postMessage({
        type: 'convert',
        requestId: 1,
        bytes,
        from: input.from,
        to: input.to,
        maxOutputBytes: input.maxOutputBytes,
      }, [bytes]);
    } catch {
      fail('failed');
    }
  });
}

function exactTransferableInput(bytes: Uint8Array): ArrayBuffer {
  if (
    bytes.byteOffset === 0
    && bytes.byteLength === bytes.buffer.byteLength
    && bytes.buffer instanceof ArrayBuffer
  ) return bytes.buffer;
  return copyToArrayBuffer(bytes);
}

function workerOutputBuffer(
  bytes: Uint8Array | ArrayBuffer,
  maxOutputBytes: number,
  from: LegacyOfficeFormat,
  to: OoxmlFormat,
): ArrayBuffer {
  if (bytes.byteLength > maxOutputBytes) {
    throw new LegacyOfficeConversionError('output-too-large', from, to);
  }
  if (bytes instanceof ArrayBuffer) return bytes;
  if (
    bytes.byteOffset === 0
    && bytes.byteLength === bytes.buffer.byteLength
    && bytes.buffer instanceof ArrayBuffer
  ) return bytes.buffer;
  return copyToArrayBuffer(bytes);
}

function copyToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  // Node.js Buffer overrides `slice()` with view semantics. An explicit copy is
  // required to avoid transferring its pooled/over-long backing allocation.
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

function isWorkerRequest(value: unknown): value is LegacyOfficeWorkerRequest {
  if (!isResponseObject(value)) return false;
  if (
    value.type !== 'convert'
    || value.requestId !== 1
    || !(value.bytes instanceof ArrayBuffer)
    || typeof value.maxOutputBytes !== 'number'
    || !Number.isSafeInteger(value.maxOutputBytes)
    || value.maxOutputBytes <= 0
    || value.maxOutputBytes > HARD_MAX_LEGACY_CONVERSION_BYTES
  ) return false;
  return (value.from === 'doc' && value.to === 'docx')
    || (value.from === 'xls' && value.to === 'xlsx')
    || (value.from === 'ppt' && value.to === 'pptx');
}

function isResponseObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object';
}

function conversionFailure(
  reason: LegacyOfficeConversionFailureReason,
  input: Pick<LegacyOfficeConversionInput, 'from' | 'to'>,
): LegacyOfficeConversionError {
  return new LegacyOfficeConversionError(reason, input.from, input.to);
}

function boundedInteger(
  value: number | undefined,
  fallback: number,
  minimum: number,
  maximum: number,
  name: string,
): number {
  const resolved = value ?? fallback;
  if (!Number.isSafeInteger(resolved) || resolved < minimum || resolved > maximum) {
    throw new TypeError(`${name} must be an integer from ${minimum} through ${maximum}`);
  }
  return resolved;
}
