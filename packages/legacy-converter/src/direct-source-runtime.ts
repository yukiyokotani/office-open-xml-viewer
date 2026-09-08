import {
  detachWasmBindgenResource,
  isWasmTrap,
  WasmTrapError,
} from '@silurus/ooxml-core/worker';

interface RuntimeHandle { invalidate(): void }
interface TrapDomain { failure?: WasmTrapError; live: Set<RuntimeHandle> }
const trapDomains = new WeakMap<object, TrapDomain>();

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
  let initializationPoison: WasmTrapError | undefined;
  let domain: TrapDomain | undefined;
  const failure = (): WasmTrapError | undefined => initializationPoison ?? domain?.failure;
  const failTrap = (error: unknown): never => {
    if (!isWasmTrap(error)) throw error;
    const normalized = failure() ?? new WasmTrapError(`${config.label} WASM runtime trapped and is unavailable`);
    if (domain) {
      domain.failure ??= normalized;
      for (const handle of domain.live) handle.invalidate();
      domain.live.clear();
    } else initializationPoison = normalized;
    throw normalized;
  };
  const runNative = <T>(operation: () => T): T => {
    const poisoned = failure();
    if (poisoned) throw poisoned;
    try { return operation(); } catch (error) { return failTrap(error); }
  };
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
      domain = trapDomains.get(glue as object);
      if (!domain) {
        domain = { live: new Set() };
        trapDomains.set(glue as object, domain);
      }
      if (domain.failure) throw domain.failure;
      try {
        const wasm = await config.resolveWasm(wasmUrl);
        await glue.default({ module_or_path: wasm });
      } catch (error) {
        return failTrap(error);
      }
      return glue;
    });
    return initialized;
  };

  return Object.freeze({
    async open(bytes, descriptor, signal) {
      const poisoned = failure();
      if (poisoned) throw poisoned;
      const validated = config.validate(descriptor);
      if (bytes.byteLength > config.maximumSourceBytes) {
        throw new RangeError(`${config.label} direct source byte budget exceeded`);
      }
      throwIfAborted(signal, config.label);
      let glue: G;
      try {
        glue = await waitForInitialization(initialize(validated.wasmUrl), signal, config.label);
      } catch (error) {
        return failTrap(error);
      }
      domain = trapDomains.get(glue as object);
      if (!domain) throw new Error(`${config.label} source runtime identity is unavailable`);
      if (domain.failure) throw domain.failure;
      throwIfAborted(signal, config.label);
      let archive: A | undefined;
      try {
        archive = runNative(() => config.construct(glue, bytes));
        throwIfAborted(signal, config.label);
      } catch (error) {
        if (archive) {
          try { closeAndFree(archive, config.closeNative, runNative); } catch {}
        }
        throw error;
      }
      let closed = false;
      let raw: A | undefined = archive;
      const handle = {
        invalidate() {
          const current = raw;
          raw = undefined;
          detachWasmBindgenResource(current);
        },
      };
      domain.live.add(handle);
      const guarded = new Proxy(archive, {
        get(_target, property) {
          const current = raw;
          if (!current) throw failure() ?? new Error(`${config.label} source archive is closed`);
          const value = runNative(() => Reflect.get(current, property, current));
          if (typeof value !== 'function') return value;
          return (...args: unknown[]) => {
            const live = raw;
            if (!live) throw failure() ?? new Error(`${config.label} source archive is closed`);
            return runNative(() => {
              const method = Reflect.get(live, property, live) as (...values: unknown[]) => unknown;
              return Reflect.apply(method, live, args);
            });
          };
        },
      }) as A;
      return {
        archive: guarded,
        sourceByteLength: bytes.byteLength,
        closeArchive() {
          if (closed) return;
          closed = true;
          domain?.live.delete(handle);
          const current = raw;
          raw = undefined;
          if (current) closeAndFree(current, config.closeNative, runNative);
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

function closeAndFree<A extends DirectSourceArchive>(archive: A, closeNative: (archive: A) => void, run: <T>(operation: () => T) => T): void {
  let primary: unknown;
  let failed = false;
  try { run(() => closeNative(archive)); } catch (error) { failed = true; primary = error; }
  // A trap invalidates the instance; never re-enter it through a destructor.
  if (isWasmTrap(primary) || primary instanceof WasmTrapError) {
    detachWasmBindgenResource(archive);
  } else {
    try { run(() => archive.free()); } catch (error) {
      if (!failed) { failed = true; primary = error; }
      if (isWasmTrap(error) || error instanceof WasmTrapError) detachWasmBindgenResource(archive);
    }
  }
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
