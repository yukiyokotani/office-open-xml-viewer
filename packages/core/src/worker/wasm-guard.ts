import { withWasmResourceErrors } from './error-wire.js';
import { RuntimeGeneration } from '../internal/runtime-generation.js';

export type WasmTrapErrorCode = 'parser-crashed';

/** A trap-shaped parser failure. The next operation uses a fresh runtime. */
export class WasmTrapError extends Error {
  readonly code: WasmTrapErrorCode = 'parser-crashed';

  constructor(message: string) {
    super(message);
    this.name = 'WasmTrapError';
    Object.setPrototypeOf(this, WasmTrapError.prototype);
  }
}

/**
 * Classify by boundary error type/name, never by message substring.
 *
 * With `panic = "abort"`, a Rust panic, an aborting allocation failure or `unreachable`
 * ends the export in a trap without a typed `Result`, and engines may report
 * OOM or stack exhaustion through implementation-defined errors (for example
 * `RangeError`, `InternalError` or `OOMError`). The cause is not preserved, so
 * `WasmTrapError` has the single code `parser-crashed` and is not relabelled
 * as out-of-memory. A plain `Error` is not treated as a trap: it cannot be
 * told apart from a graceful parser `Result::Err`.
 */
export function isWasmTrap(error: unknown): boolean {
  const RuntimeError = (globalThis as { WebAssembly?: { RuntimeError?: unknown } }).WebAssembly
    ?.RuntimeError as (new () => Error) | undefined;
  if (RuntimeError && error instanceof RuntimeError) return true;
  if (error instanceof RangeError) return true;
  if (!(error instanceof Error)) return false;
  return error.name === 'RuntimeError'
    || error.name === 'CompileError'
    || error.name === 'LinkError'
    || error.name === 'InternalError'
    || error.name === 'OOMError';
}

/**
 * Remove a wasm-bindgen object's GC finalizer without invoking its Rust
 * destructor. The generated `__destroy_into_raw()` method only clears the JS
 * pointer and unregisters the wrapper from its `FinalizationRegistry`; the
 * discarded runtime generation continues to own the underlying allocation.
 *
 * This must run before `reinit()` replaces wasm-bindgen's module-level `wasm`
 * binding. Otherwise a stale wrapper collected later calls its finalizer with
 * an old-generation pointer against the new instance's linear memory.
 */
export function detachWasmBindgenResource(resource: unknown): void {
  try {
    if ((typeof resource !== 'object' || resource === null) && typeof resource !== 'function') return;
    const detach = Reflect.get(resource as object, '__destroy_into_raw');
    if (typeof detach === 'function') Reflect.apply(detach, resource, []);
  } catch {
    // Poison fan-out must still invalidate every sibling and preserve the
    // original trap. Production wasm-bindgen wrappers use a JS-only method here;
    // focused integration coverage pins that generated contract.
  }
}

export type WasmInitInput =
  | string
  | URL
  | Request
  | ArrayBuffer
  | ArrayBufferView
  | WebAssembly.Module
  | Response;

export interface WasmInitOptions {
  readonly module_or_path: WasmInitInput | Promise<WasmInitInput>;
}

export type WasmInit = (options: WasmInitOptions) => Promise<unknown>;
export type WasmReinit = (options: WasmInitOptions) => Promise<unknown>;

function invokeWasmInitializer(
  initializer: WasmInit | WasmReinit,
  input: WasmInitInput,
): Promise<unknown> {
  return initializer({ module_or_path: input });
}

export interface WasmParserHostOptions<TArchive> {
  readonly freeArchive?: (archive: TArchive) => void;
  /** Production callers provide the wasm-bindgen `reinit` export. */
  readonly reinit?: WasmReinit;
}

/**
 * Worker-facing single-archive facade over the same RuntimeGeneration used by
 * format-owned in-process sessions. It owns the archive reference; the shared
 * primitive alone owns initialization, poison fan-out, and recovery ordering.
 */
export class WasmParserHost<TArchive = unknown> {
  private readonly runtime: RuntimeGeneration<WasmTrapError>;
  private wasmInput: WasmInitInput | null = null;
  private currentArchive: TArchive | null = null;

  constructor(
    private readonly init: WasmInit,
    private readonly options: WasmParserHostOptions<TArchive> = {},
  ) {
    this.runtime = new RuntimeGeneration(
      () => this.invokeConfigured(this.init),
      () => this.invokeConfigured(this.options.reinit ?? this.init),
      normalizeTrap,
    );
    this.runtime.onPoison(() => this.dropPoisonedArchive());
  }

  setWasmInput(input: WasmInitInput): void {
    this.wasmInput = input;
    // Preserve the historical eager first initialization. The rejection is
    // observed by ensureReady() on the request path.
    this.runtime.ensureReady().catch(() => undefined);
  }

  /** @deprecated Use setWasmInput. */
  setWasmUrl(input: WasmInitInput): void {
    this.setWasmInput(input);
  }

  get archive(): TArchive | null {
    return this.currentArchive;
  }

  setArchive(archive: TArchive): void {
    this.freeArchive();
    this.currentArchive = archive;
  }

  disposeArchive(): void {
    this.freeArchive();
  }

  get poisoned(): boolean {
    return this.runtime.poisoned;
  }

  async ensureReady(): Promise<void> {
    await this.runtime.ensureReady();
  }

  run<TResult>(operation: () => TResult): TResult {
    return withWasmResourceErrors(() => this.runtime.run(operation));
  }

  poison(): void {
    this.runtime.poison(new WasmTrapError('WASM parser was recycled'));
  }

  private invokeConfigured(initializer: WasmInit | WasmReinit): Promise<unknown> {
    if (this.wasmInput === null) {
      return Promise.reject(new Error('WasmParserHost: setWasmInput was never called'));
    }
    return invokeWasmInitializer(initializer, this.wasmInput);
  }

  private freeArchive(): void {
    if (this.currentArchive !== null && this.options.freeArchive) {
      this.options.freeArchive(this.currentArchive);
    }
    this.currentArchive = null;
  }

  private dropPoisonedArchive(): void {
    // A trap invalidates the entire runtime generation. Calling a wasm-bindgen
    // destructor here would re-enter poisoned linear memory and may trap again;
    // detach the JS handle and let the discarded instance be collected whole.
    const archive = this.currentArchive;
    this.currentArchive = null;
    detachWasmBindgenResource(archive);
  }
}

function normalizeTrap(error: unknown): WasmTrapError | null {
  if (!isWasmTrap(error)) return null;
  const detail = error instanceof Error ? error.message : String(error);
  return new WasmTrapError(`WASM parser trapped and was recycled: ${detail}`);
}
