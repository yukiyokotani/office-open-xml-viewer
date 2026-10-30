import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from '@silurus/ooxml-core/internal/model-source-runtime';
import { odfNativeRejection } from './native-rejection.js';

/** The DOCX model-source archive the ODT reader exposes. */
export interface OdtNativeDocument {
  free(): void;
  open_document_cursor(operationId: number, generation: number): void;
  pull_document_chunk(
    sequence: number,
    operationId: number,
    generation: number,
    byteCredit: number,
  ): Uint8Array;
  document_chunk_done(): boolean;
  acknowledge_document_chunk(sequence: number, operationId: number, generation: number): void;
  cancel_document_cursor(): void;
  close_document_session(): void;
  assert_healthy(): void;
  extract_image(key: string): Uint8Array;
  image_mime_type(key: string): string;
}

export type OwnedOdtSource = OwnedDirectSource<OdtNativeDocument>;

export interface OdtGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  OdtDocument: new (
    bytes: Uint8Array,
    maxArchiveEntryBytes?: number,
    maxTotalInflatedBytes?: number,
    maxArchiveEntries?: number,
  ) => OdtNativeDocument;
}

/** Internal injectable engine; generated glue is wired only by the source module. */
export function createOdtSourceEngine(
  loadGlue: () => Promise<OdtGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
): Readonly<{
  open(
    bytes: Uint8Array,
    wasmUrl: string,
    context: ModelSourceOpenContext | undefined,
    signal?: AbortSignal,
  ): Promise<OwnedOdtSource>;
}> {
  const runtime = createDirectSourceRuntime<OdtGlue, OdtNativeDocument, ModelSourceOpenContext>({
    label: 'OpenDocument ODT',
    mapNativeError: (error) => odfNativeRejection(error, 'odt'),
    loadGlue,
    resolveWasm,
    // `null` (a caller-removed limit) and an absent context both leave the
    // reader's hard ceilings as the only bound.
    construct: (glue, bytes, context) => new glue.OdtDocument(
      bytes,
      context?.resourcePolicy.maxArchiveEntryBytes ?? undefined,
      context?.resourcePolicy.maxTotalInflatedBytes ?? undefined,
      context?.resourcePolicy.maxArchiveEntries ?? undefined,
    ),
    closeNative: (document) => document.close_document_session(),
  });
  return Object.freeze({
    open: (bytes, wasmUrl, context, signal) => runtime.open(bytes, wasmUrl, signal, context),
  });
}

export { resolveDirectWasmInput };
