import { readNativePictureBulletPolicy, type NativePictureBulletPolicy } from './doc-picture-bullet-policy.js';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from './direct-source-runtime.js';

import type { NativeWordBreakingPolicy } from './legacy-doc.js';
import { readNativeDocReadingPolicies, type NativeDocReadingPolicies } from './source-module-config.js';

/** Single-capability strings remain accepted at this internal engine boundary.
 * Clone-safe source descriptors pass independent policies together. */
type LegacyDocReadingRequest = NativeDocReadingPolicies | NativeWordBreakingPolicy | 'relocateForReading';

import { MAX_LEGACY_SOURCE_BYTES as MAX_LEGACY_DOC_SOURCE_BYTES } from './legacy-source-limits.js';

export { MAX_LEGACY_DOC_SOURCE_BYTES };

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
  /** MS-DOC 2.7.2 DopBase fRMPrint: the DOC shows its revision markup in
   *  print/PDF output (Word's PDF is the display target for legacy DOC). */
  revision_markup_in_print?(): boolean;
  /** MS-DOC 2.7.2 DopBase fRMView: the DOC shows its revision markup on screen. */
  revision_markup_on_screen?(): boolean;
  /** Whether the projected model carries any revision mark. */
  has_revision_marks?(): boolean;
  native_picture_bullet_reading_requested?(): boolean;
}

export type OwnedLegacyDocSource = OwnedDirectSource<LegacyDocNativeDocument>;

export interface LegacyDocGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  LegacyDocDocument: new (bytes: Uint8Array, modelBudget?: number, relocateForReading?: boolean,
    simplifyWordBreakingForReading?: boolean, storedPictureBulletsForReading?: boolean) => LegacyDocNativeDocument;
}

/** Internal injectable engine; generated glue is wired only after its facade is accepted. */
export function createLegacyDocSourceEngine(
  loadGlue: () => Promise<LegacyDocGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
  modelBudget?: number,
  pictureBulletPolicy?: NativePictureBulletPolicy,
): Readonly<{
  open(bytes: Uint8Array, wasmUrl: string, signal?: AbortSignal, request?: LegacyDocReadingRequest): Promise<OwnedLegacyDocSource>;
}> {
  const picturePolicy = readNativePictureBulletPolicy(pictureBulletPolicy);
  if (modelBudget !== undefined && (
    !Number.isSafeInteger(modelBudget) || modelBudget <= 0
    || modelBudget > MAX_LEGACY_DOC_SOURCE_BYTES
  )) {
    throw new RangeError('legacy DOC model budget is invalid');
  }
  return createDirectSourceRuntime<LegacyDocGlue, LegacyDocNativeDocument, LegacyDocReadingRequest>({
    label: 'legacy DOC',
    maximumSourceBytes: MAX_LEGACY_DOC_SOURCE_BYTES,
    loadGlue,
    resolveWasm,
    construct: (glue, bytes, request) => {
      let selected: NativeDocReadingPolicies;
      if (request === undefined || request === 'strict') selected = {};
      else if (request === 'relocateForReading') selected = { nativeContourPolicy: request };
      else if (request === 'simplifyForReading') selected = { nativeWordBreakingPolicy: request };
      else if (typeof request === 'object' && request !== null && !Array.isArray(request)) selected = request;
      else throw new TypeError('Invalid native DOC reading policy');
      const policies = readNativeDocReadingPolicies({ ...selected,
        nativePictureBulletPolicy: selected.nativePictureBulletPolicy === undefined
          ? picturePolicy : selected.nativePictureBulletPolicy });
      const contour = policies.nativeContourPolicy === 'relocateForReading';
      const wordBreaking = policies.nativeWordBreakingPolicy === 'simplifyForReading';
      const pictureBullets = policies.nativePictureBulletPolicy === 'storedSizeForReading';
      // Native ABI slots are contour=3, word breaking=4, picture bullets=5.
      // The all-strict path preserves the default native constructor behavior.
      return contour || wordBreaking || pictureBullets
        ? new glue.LegacyDocDocument(bytes, modelBudget, contour, wordBreaking, pictureBullets)
        : new glue.LegacyDocDocument(bytes, modelBudget);
    },
    // Closing ends the pull cursor. free()/Drop owns the retained image resources.
    closeNative: (document) => document.close_document_session(),
  });
}

/**
 * The DOCX view the DOC asks for: its revision markup when it prints that
 * markup (fRMPrint) and actually carries revision marks. Word's PDF output is
 * the display target for legacy DOC, so the print setting decides.
 */
export function legacyDocViewDefaults(
  document: Pick<LegacyDocNativeDocument, 'revision_markup_in_print' | 'has_revision_marks'>,
): Readonly<{ showTrackedChanges?: boolean }> {
  return document.revision_markup_in_print?.() === true && document.has_revision_marks?.() === true
    ? { showTrackedChanges: true }
    : {};
}

export { resolveDirectWasmInput };
