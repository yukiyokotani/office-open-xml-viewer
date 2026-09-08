import {
  validateLegacyDocSourceDescriptor,
  MAX_LEGACY_DOC_SOURCE_BYTES,
  type LegacyDocDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-doc-source';
import {
  createDirectSourceRuntime,
  type OwnedDirectSource,
} from './direct-source-runtime.js';

export interface LegacyDocNativeDocument {
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
  image_mime_type?(key: string): string;
}

export type OwnedLegacyDocSource = OwnedDirectSource<LegacyDocNativeDocument>;

export interface LegacyDocGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  LegacyDocDocument: new (bytes: Uint8Array, modelBudget?: number) => LegacyDocNativeDocument;
}

/** Internal injectable engine; generated glue is wired only after its facade is accepted. */
export function createLegacyDocSourceEngine(
  loadGlue: () => Promise<LegacyDocGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
  modelBudget?: number,
): Readonly<{
  open(bytes: Uint8Array, descriptor: LegacyDocDirectSourceDescriptor, signal?: AbortSignal): Promise<OwnedLegacyDocSource>;
}> {
  if (modelBudget !== undefined && (
    !Number.isSafeInteger(modelBudget) || modelBudget <= 0
    || modelBudget > MAX_LEGACY_DOC_SOURCE_BYTES
  )) {
    throw new RangeError('legacy DOC model budget is invalid');
  }
  return createDirectSourceRuntime({
    label: 'legacy DOC',
    maximumSourceBytes: MAX_LEGACY_DOC_SOURCE_BYTES,
    validate: validateLegacyDocSourceDescriptor,
    loadGlue,
    resolveWasm,
    construct: (glue, bytes) => new glue.LegacyDocDocument(bytes, modelBudget),
    // Closing ends the pull cursor. free()/Drop owns the retained image resources.
    closeNative: (document) => document.close_document_session(),
  });
}
