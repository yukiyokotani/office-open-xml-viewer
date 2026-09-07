export interface DirectSourceDescriptor {
  readonly wasmUrl: string;
}

export interface DirectSourceArchive {
  free(): void;
}

export interface OwnedDirectSource<A> {
  readonly archive: A;
  readonly sourceByteLength: number;
  closeArchive(): void;
}

export interface DirectSourceGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
}

export function createDirectSourceRuntime<D extends DirectSourceDescriptor, G extends DirectSourceGlue, A extends DirectSourceArchive>(
  config: Readonly<{
    label: string;
    maximumSourceBytes: number;
    validate(value: unknown): D;
    loadGlue(): Promise<G>;
    resolveWasm(wasmUrl: string): Promise<unknown>;
    construct(glue: G, bytes: Uint8Array): A;
    closeNative(archive: A): void;
  }>,
): Readonly<{ open(bytes: Uint8Array, descriptor: D, signal?: AbortSignal): Promise<OwnedDirectSource<A>> }> {
  let initialized: Promise<G> | undefined;
  let initializedUrl: string | undefined;
  const initialize = (wasmUrl: string): Promise<G> => {
    if (initialized) {
      if (wasmUrl !== initializedUrl) {
        return Promise.reject(new Error(`${config.label} source engine is pinned to another WASM URL`));
      }
      return initialized;
    }
    initializedUrl = wasmUrl;
    // Generated glue is realm-global. Initialization failure may leave a
    // partially installed instance, so never retry or replace its generation
    // while native sessions may still refer to it.
    initialized = config.loadGlue().then(async (glue) => {
      const wasm = await config.resolveWasm(wasmUrl);
      await glue.default({ module_or_path: wasm });
      return glue;
    });
    return initialized;
  };

  return Object.freeze({
    async open(bytes, descriptor, signal) {
      const validated = config.validate(descriptor);
      if (bytes.byteLength > config.maximumSourceBytes) {
        throw new RangeError(`${config.label} direct source byte budget exceeded`);
      }
      throwIfAborted(signal, config.label);
      const glue = await waitForInitialization(initialize(validated.wasmUrl), signal, config.label);
      throwIfAborted(signal, config.label);
      let archive: A | undefined;
      try {
        archive = config.construct(glue, bytes);
        throwIfAborted(signal, config.label);
      } catch (error) {
        if (archive) {
          try { closeAndFree(archive, config.closeNative); } catch {}
        }
        throw error;
      }
      let closed = false;
      return {
        archive,
        sourceByteLength: bytes.byteLength,
        closeArchive() {
          if (closed) return;
          closed = true;
          closeAndFree(archive, config.closeNative);
        },
      };
    },
  });
}

export async function resolveDirectWasmInput(wasmUrl: string): Promise<unknown> {
  const url = new URL(wasmUrl, import.meta.url);
  const nodeProcess = (globalThis as { process?: { versions?: { node?: string } } }).process;
  if (url.protocol === 'file:' && nodeProcess?.versions?.node) {
    const nodeFsPromises: string = 'node:fs/promises';
    const { readFile } = await import(/* @vite-ignore */ nodeFsPromises);
    return readFile(url);
  }
  return url;
}

function closeAndFree<A extends DirectSourceArchive>(archive: A, closeNative: (archive: A) => void): void {
  let primary: unknown;
  let failed = false;
  try { closeNative(archive); } catch (error) { failed = true; primary = error; }
  try { archive.free(); } catch (error) { if (!failed) { failed = true; primary = error; } }
  if (failed) throw primary;
}

function waitForInitialization<T>(pending: Promise<T>, signal: AbortSignal | undefined, label: string): Promise<T> {
  if (!signal) return pending;
  throwIfAborted(signal, label);
  return new Promise<T>((resolve, reject) => {
    const abort = (): void => { cleanup(); try { throwIfAborted(signal, label); } catch (error) { reject(error); } };
    const cleanup = (): void => signal.removeEventListener('abort', abort);
    signal.addEventListener('abort', abort, { once: true });
    pending.then((value) => { cleanup(); resolve(value); }, (error: unknown) => { cleanup(); reject(error); });
  });
}

function throwIfAborted(signal: AbortSignal | undefined, label: string): void {
  if (!signal?.aborted) return;
  const error = new Error(`${label} direct source was aborted`);
  error.name = 'AbortError';
  throw error;
}
