import type { LegacyOfficeConversionOptions } from '@silurus/ooxml-core';
import type { PptxPresentation } from './presentation.js';

interface BoundLegacyConversion {
  readonly options?: LegacyOfficeConversionOptions;
  readonly cleanup: () => void;
}

/**
 * Keep a direct PPT source's combined cancellation signal wired for the owned
 * presentation lifetime. Byte converters need their temporary signal wiring
 * only until conversion/load settles, preserving their existing cleanup scope.
 */
export function settleLegacyPptLoad(
  pending: Promise<PptxPresentation>,
  conversion: BoundLegacyConversion,
): Promise<PptxPresentation> {
  const selected = conversion.options?.ppt;
  if (selected === undefined) return pending;
  if (!('source' in selected)) return pending.finally(conversion.cleanup);
  return pending.then(
    (presentation) => {
      presentation._retainLegacyPptSignalCleanup(conversion.cleanup);
      return presentation;
    },
    (error: unknown) => {
      conversion.cleanup();
      throw error;
    },
  );
}
