import type { OoxmlFormat } from '../errors/ooxml-error.js';

export type LegacyOfficeFormat = 'doc' | 'xls' | 'ppt';

export type LegacyOfficeConversionFailureReason =
  | 'aborted'
  | 'timeout'
  | 'source-too-large'
  | 'output-too-large'
  | 'capacity-exceeded'
  | 'unsupported-input'
  | 'failed'
  | 'invalid-output';

/** Stable typed failure for the opt-in conversion stage. */
export class LegacyOfficeConversionError extends Error {
  readonly code = 'legacy-office-conversion' as const;
  readonly stage = 'conversion' as const;
  readonly reason: LegacyOfficeConversionFailureReason;
  readonly from: LegacyOfficeFormat;
  readonly to: OoxmlFormat;

  constructor(
    reason: LegacyOfficeConversionFailureReason,
    from: LegacyOfficeFormat,
    to: OoxmlFormat,
    message = conversionErrorMessage(reason),
  ) {
    super(message);
    this.name = 'LegacyOfficeConversionError';
    this.reason = reason;
    this.from = from;
    this.to = to;
    Object.setPrototypeOf(this, LegacyOfficeConversionError.prototype);
  }
}

function conversionErrorMessage(reason: LegacyOfficeConversionFailureReason): string {
  switch (reason) {
    case 'aborted':
      return 'Legacy Office conversion was aborted.';
    case 'timeout':
      return 'Legacy Office conversion exceeded its execution time limit.';
    case 'source-too-large':
      return 'The legacy Office input exceeds the configured conversion size limit.';
    case 'output-too-large':
      return 'The converted OOXML package exceeds the configured conversion size limit.';
    case 'capacity-exceeded':
      return 'The legacy Office converter has reached its bounded queue capacity.';
    case 'unsupported-input':
      return 'The converter does not support this legacy Office input.';
    case 'failed':
      return 'The legacy Office converter failed.';
    case 'invalid-output':
      return 'The converter did not produce a valid macro-free OOXML package for the requested format.';
    default:
      reason satisfies never;
      return 'The legacy Office converter failed.';
  }
}
