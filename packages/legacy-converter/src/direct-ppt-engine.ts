import {
  validateLegacyPptSourceDescriptor,
  type LegacyPptDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-ppt-source';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from './direct-source-runtime.js';

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

export type OwnedLegacyPptSource = OwnedDirectSource<LegacyPptNativeArchive>;

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
  return createDirectSourceRuntime({
    label: 'legacy PPT',
    maximumSourceBytes: MAX_DIRECT_PPT_SOURCE_BYTES,
    validate: validateLegacyPptSourceDescriptor,
    loadGlue,
    resolveWasm,
    construct: (glue, bytes) => new glue.LegacyPptPresentation(bytes),
    closeNative: (archive) => archive.close_presentation_session(),
  });
}

// Generated glue is a realm singleton. Keep exactly one production engine and
// pin it to the first attempted asset URL in this realm, including sticky failure.
const defaultEngine = createLegacyPptSourceEngine(
  () => import('./wasm-direct-ppt/legacy_office_converter.js'),
  resolveDirectWasmInput,
);

export function openLegacyPptSource(
  bytes: Uint8Array,
  descriptor: LegacyPptDirectSourceDescriptor,
  signal?: AbortSignal,
): Promise<OwnedLegacyPptSource> {
  return defaultEngine.open(bytes, descriptor, signal);
}
