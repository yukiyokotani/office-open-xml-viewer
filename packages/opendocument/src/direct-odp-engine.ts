import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from '@silurus/ooxml-core/internal/model-source-runtime';
import { odfNativeRejection } from './native-rejection.js';

/** The PPTX model-source archive the ODP reader exposes. */
export interface OdpNativePresentation {
  free(): void;
  presentation_bootstrap(): Uint8Array;
  pull_slide(slideIndex: number, operationId: number, generation: number, byteCredit: number): Uint8Array;
  acknowledge_slide(operationId: number, generation: number): void;
  cancel_slide(): void;
  close_presentation_session(): void;
  assert_healthy(): void;
  extract_image(path: string): Uint8Array;
  image_mime_type(path: string): string;
  slide_cursor_resource_usage(): Uint8Array;
}

export type OwnedOdpSource = OwnedDirectSource<OdpNativePresentation>;

export interface OdpGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  OdpPresentation: new (
    bytes: Uint8Array,
    maxArchiveEntryBytes?: number,
    maxTotalInflatedBytes?: number,
    maxArchiveEntries?: number,
  ) => OdpNativePresentation;
}

/** Internal injectable engine; generated glue is wired only by the source module. */
export function createOdpSourceEngine(
  loadGlue: () => Promise<OdpGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
): Readonly<{
  open(
    bytes: Uint8Array,
    wasmUrl: string,
    context: ModelSourceOpenContext | undefined,
    signal?: AbortSignal,
  ): Promise<OwnedOdpSource>;
}> {
  const runtime = createDirectSourceRuntime<OdpGlue, OdpNativePresentation, ModelSourceOpenContext>({
    label: 'OpenDocument ODP',
    mapNativeError: (error) => odfNativeRejection(error, 'odp'),
    loadGlue,
    resolveWasm,
    // The shared governor's WASM encoding: absent → standard limit, 0 → the
    // caller removed the limit (hard ceilings still apply).
    construct: (glue, bytes, context) => new glue.OdpPresentation(
      bytes,
      wasmLimit(context?.resourcePolicy.maxArchiveEntryBytes),
      wasmLimit(context?.resourcePolicy.maxTotalInflatedBytes),
      wasmLimit(context?.resourcePolicy.maxArchiveEntries),
    ),
    closeNative: (presentation) => presentation.close_presentation_session(),
  });
  return Object.freeze({
    open: (bytes, wasmUrl, context, signal) => runtime.open(bytes, wasmUrl, signal, context),
  });
}

function wasmLimit(value: number | null | undefined): number | undefined {
  return value === null ? 0 : value;
}

export { resolveDirectWasmInput };
