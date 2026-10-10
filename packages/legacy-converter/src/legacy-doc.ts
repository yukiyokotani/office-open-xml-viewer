import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, type LegacySourceOptions } from './legacy-source.js';
import { readNativeDocReadingPolicies, type NativeDocReadingPolicies } from './source-module-config.js';
import defaultWasmUrl from './wasm-direct-doc/legacy_doc_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-doc-source-module.ts?url';

/** Explicit DOC reading capability; strict is the default and retains operand
 * domain validation. `simplifyForReading` preserves literal logical text and
 * correctly framed two-byte raw Hresi, including values outside the normative
 * Word domain. It does not interpret or perform Hresi transformations.
 * Retained reading pages carry a disclosure; built-in viewers visibly warn
 * that line and page breaks may differ. */
export type NativeWordBreakingPolicy = 'strict' | 'simplifyForReading';
export type { NativePictureBulletPolicy } from './doc-picture-bullet-policy.js';
export interface LegacyDocSourceOptions extends LegacySourceOptions, NativeDocReadingPolicies {
  /** Explicitly relocate complete drawings into reading blocks; automatic
   * Office Tight/Through contours are not claimed. */
  readonly nativeContourPolicy?: 'strict' | 'relocateForReading';
  /** Preserve logical text while allowing line and page breaks to differ. */
  readonly nativeWordBreakingPolicy?: NativeWordBreakingPolicy;
  /** Use stored embedded picture sizes; successful visible markers disclose
   * this reading choice rather than an Office AUTO sizing claim. */
  readonly nativePictureBulletPolicy?: 'strict' | 'storedSizeForReading';
}

/**
 * A model source for legacy Word 97-2003 (.doc) input, for
 * `LoadOptions.modelSources` of the DOCX loaders and viewers. Creating it
 * fetches nothing; the source module and its WASM load only when a .doc input
 * is claimed.
 */
export function legacyDocSource(options: LegacyDocSourceOptions = {}): ModelSource<'docx'> {
  if (typeof options !== 'object' || options === null || Array.isArray(options))
    throw new TypeError('legacy DOC source options must be an object');
  const policies = readNativeDocReadingPolicies({
    nativeContourPolicy: options.nativeContourPolicy,
    nativeWordBreakingPolicy: options.nativeWordBreakingPolicy,
    nativePictureBulletPolicy: options.nativePictureBulletPolicy,
  });
  return createLegacySource<'docx'>('doc', options, {
    wasmUrl: new URL(defaultWasmUrl, import.meta.url).href,
    moduleUrl: new URL(defaultModuleUrl, import.meta.url).href,
  }, {
    ...(policies.nativeContourPolicy === 'strict' ? {} : { nativeContourPolicy: policies.nativeContourPolicy }),
    ...(policies.nativeWordBreakingPolicy === 'strict' ? {} : { nativeWordBreakingPolicy: policies.nativeWordBreakingPolicy }),
    ...(policies.nativePictureBulletPolicy === 'strict' ? {} : { nativePictureBulletPolicy: policies.nativePictureBulletPolicy }),
  });
}
