import { acquirePptxNodeSession } from '@silurus/ooxml-pptx/internal/session';
import type { PptxNodeArchive } from '@silurus/ooxml-pptx/internal/session';
import { resolveCjkFallback } from '@silurus/ooxml-core';
import type { acquirePptxSessionFromArchive } from '@silurus/ooxml-pptx/internal/model-source-session';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import { resolveNodeSessionInput } from './model-source.ts';
import { PptxPresentationSessionImpl, type OpenPptxPresentationOptions } from './pptx.ts';

export type PptxModelSourceAcquisition = ReturnType<typeof acquirePptxSessionFromArchive>;

export async function openPptxSource(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenPptxPresentationOptions,
  wasmModule: () => WebAssembly.Module,
): Promise<PptxPresentationSessionImpl> {
  const acquired = await acquirePptxInput(buffer, options, wasmModule);
  // The ordinary session's media method is required. Fill the selected
  // source's optional capability here, where its error mapping remains lazy.
  const archive = new Proxy(acquired.archive, {
    get(target, key) {
      if (key === 'extract_media' && !target.extract_media) {
        return () => { throw new Error('media extraction is unsupported for this source'); };
      }
      const value = Reflect.get(target, key);
      return typeof value === 'function' ? value.bind(target) : value;
    },
  }) as unknown as PptxNodeArchive;
  return new PptxPresentationSessionImpl(
    acquired.closeArchive, archive, acquired.bootstrap,
    acquired.metrics, options.signal, resolveCjkFallback(options.cjkFallback),
  );
}

/** Selected-source PPTX acquisition, absent from the ordinary Node entry. */
export async function acquirePptxInput(
  buffer: ArrayBuffer | Uint8Array,
  options: OoxmlNodeSessionOptions & Pick<OpenPptxPresentationOptions, 'legacyPpt' | 'odp'>,
  wasmModule: () => WebAssembly.Module,
): Promise<PptxModelSourceAcquisition> {
  const {
    acquirePptxSessionFromArchive,
    validatePptxModelSourceArchive,
    validatePptxModelSourceViewDefaults,
  } = await import('@silurus/ooxml-pptx/internal/model-source-session');
  const input = await resolveNodeSessionInput(
    buffer, 'pptx', { ...options, sources: [options.legacyPpt, options.odp] }, validatePptxModelSourceArchive,
  );
  if (input.kind === 'ooxml') {
    return acquirePptxNodeSession(input.bytes, wasmModule(), options);
  }
  try {
    validatePptxModelSourceViewDefaults(input.opened.viewDefaults);
  } catch (error) {
    try { input.opened.close(); } catch {}
    throw error;
  }
  return acquirePptxSessionFromArchive({
    archive: input.opened.archive,
    sourceByteLength: input.sourceByteLength,
    closeArchive: input.opened.close,
  }, options);
}
