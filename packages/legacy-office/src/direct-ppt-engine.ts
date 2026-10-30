import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from '@silurus/ooxml-core/internal/model-source-runtime';
import { nativeRejection } from './native-rejection.js';

export interface LegacyPptNativeArchive {
  free(): void;
  presentation_bootstrap(): Uint8Array;
  pull_slide(slideIndex: number, operationId: number, generation: number, byteCredit: number): Uint8Array;
  acknowledge_slide(operationId: number, generation: number): void;
  cancel_slide(): void;
  close_presentation_session(): void;
  assert_healthy(): void;
  extract_image(path: string): Uint8Array;
  slide_cursor_resource_usage(): Uint8Array;
}

export type OwnedLegacyPptSource = OwnedDirectSource<LegacyPptNativeArchive>;

export interface LegacyPptGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  LegacyPptPresentation: new (bytes: Uint8Array) => LegacyPptNativeArchive;
}

/** Internal engine factory; injectable dependencies keep ownership tests content-free. */
export function createLegacyPptSourceEngine(
  loadGlue: () => Promise<LegacyPptGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
): Readonly<{
  open(bytes: Uint8Array, wasmUrl: string, signal?: AbortSignal): Promise<OwnedLegacyPptSource>;
}> {
  return createDirectSourceRuntime({
    label: 'legacy PPT',
    mapNativeError: (error) => nativeRejection(error, 'ppt'),
    loadGlue,
    resolveWasm,
    construct: (glue, bytes) => new glue.LegacyPptPresentation(bytes),
    closeNative: (archive) => archive.close_presentation_session(),
  });
}

export { resolveDirectWasmInput };
