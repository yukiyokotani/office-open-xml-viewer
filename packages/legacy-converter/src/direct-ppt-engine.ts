import {
  validateLegacyPptSourceDescriptor,
  type LegacyPptDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-ppt-source';

const MAX_DIRECT_PPT_SOURCE_BYTES = 256 * 1024 * 1024;

export interface LegacyPptNativeArchive {
  free(): void;
  presentation_bootstrap(): Uint8Array;
  pull_slide(slideIndex: number, operationId: number, generation: number, byteCredit: number): Uint8Array;
  acknowledge_slide(operationId: number, generation: number): void;
  cancel_slide(): void;
  close_presentation_session(): void;
  assert_healthy(): void;
  extract_image(path: string): Uint8Array;
  extract_media(path: string): Uint8Array;
  extract_font(path: string): Uint8Array;
  slide_cursor_resource_usage(): Uint8Array;
}

export interface OwnedLegacyPptSource {
  readonly archive: LegacyPptNativeArchive;
  readonly sourceByteLength: number;
  closeArchive(): void;
}

interface LegacyPptGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  LegacyPptPresentation: new (bytes: Uint8Array) => LegacyPptNativeArchive;
}

type LoadGlue = () => Promise<LegacyPptGlue>;
type ResolveWasm = (wasmUrl: string) => Promise<unknown>;

/** Internal engine factory; injectable dependencies keep ownership tests content-free. */
export function createLegacyPptSourceEngine(
  loadGlue: LoadGlue,
  resolveWasm: ResolveWasm,
): Readonly<{
  open(
    bytes: Uint8Array,
    descriptor: LegacyPptDirectSourceDescriptor,
    signal?: AbortSignal,
  ): Promise<OwnedLegacyPptSource>;
}> {
  let initialized: Promise<LegacyPptGlue> | undefined;
  let initializedUrl: string | undefined;

  const initialize = (wasmUrl: string): Promise<LegacyPptGlue> => {
    if (initialized) {
      if (wasmUrl !== initializedUrl) {
        return Promise.reject(new Error('legacy PPT source engine is pinned to another WASM URL'));
      }
      return initialized;
    }
    initializedUrl = wasmUrl;
    // Initialization failure is sticky. Generated glue is realm-global and can
    // fail after partially installing its instance; retrying it is not known
    // safe while this engine may own native sessions.
    initialized = loadGlue()
      .then(async (glue) => {
        const wasm = await resolveWasm(wasmUrl);
        await glue.default({ module_or_path: wasm });
        return glue;
      });
    return initialized;
  };

  return Object.freeze({
    async open(bytes, descriptor, signal) {
      const validated = validateLegacyPptSourceDescriptor(descriptor);
      if (bytes.byteLength > MAX_DIRECT_PPT_SOURCE_BYTES) {
        throw new RangeError('legacy PPT direct source byte budget exceeded');
      }
      throwIfAborted(signal);
      const glue = await waitForInitialization(initialize(validated.wasmUrl), signal);
      throwIfAborted(signal);
      let archive: LegacyPptNativeArchive | undefined;
      try {
        archive = new glue.LegacyPptPresentation(bytes);
        throwIfAborted(signal);
      } catch (error) {
        if (archive) {
          try { closeAndFree(archive); } catch {}
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
          closeAndFree(archive);
        },
      };
    },
  });
}

// Generated glue is a realm singleton. Keep exactly one production engine and
// pin it to the first attempted asset URL in this realm, including sticky failure.
const defaultEngine = createLegacyPptSourceEngine(
  () => import('./wasm-direct-ppt/legacy_office_converter.js'),
  resolveWasmInput,
);

export function openLegacyPptSource(
  bytes: Uint8Array,
  descriptor: LegacyPptDirectSourceDescriptor,
  signal?: AbortSignal,
): Promise<OwnedLegacyPptSource> {
  return defaultEngine.open(bytes, descriptor, signal);
}

async function resolveWasmInput(wasmUrl: string): Promise<unknown> {
  const url = new URL(wasmUrl, import.meta.url);
  const nodeProcess = (globalThis as { process?: { versions?: { node?: string } } }).process;
  if (url.protocol === 'file:' && nodeProcess?.versions?.node) {
    // Keep the optional Node builtin outside browser consumers' static module
    // graph and type environment. This branch is guarded by the Node runtime.
    const nodeFsPromises: string = 'node:fs/promises';
    const { readFile } = await import(/* @vite-ignore */ nodeFsPromises);
    return readFile(url);
  }
  return url;
}

function closeAndFree(archive: LegacyPptNativeArchive): void {
  let primary: unknown;
  let failed = false;
  try {
    archive.close_presentation_session();
  } catch (error) {
    failed = true;
    primary = error;
  }
  try {
    archive.free();
  } catch (error) {
    if (!failed) {
      failed = true;
      primary = error;
    }
  }
  if (failed) throw primary;
}

function waitForInitialization<T>(pending: Promise<T>, signal: AbortSignal | undefined): Promise<T> {
  if (!signal) return pending;
  throwIfAborted(signal);
  return new Promise<T>((resolve, reject) => {
    const abort = (): void => {
      cleanup();
      try {
        throwIfAborted(signal);
      } catch (error) {
        reject(error);
      }
    };
    const cleanup = (): void => signal.removeEventListener('abort', abort);
    signal.addEventListener('abort', abort, { once: true });
    pending.then(
      (value) => {
        cleanup();
        resolve(value);
      },
      (error: unknown) => {
        cleanup();
        reject(error);
      },
    );
  });
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('legacy PPT direct source was aborted');
  error.name = 'AbortError';
  throw error;
}
