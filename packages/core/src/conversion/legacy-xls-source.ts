import { validateLegacySourceDescriptor, type LegacyDirectSourceDescriptor } from './legacy-source-descriptor.js';

/** First-party native decoder selection for XLS. */
export interface LegacyXlsDirectSourceDescriptor extends LegacyDirectSourceDescriptor<'xls'> {}

export const MAX_LEGACY_XLS_SOURCE_BYTES = 256 * 1024 * 1024;

export function validateLegacyXlsSourceDescriptor(value: unknown): LegacyXlsDirectSourceDescriptor {
  return validateLegacySourceDescriptor(value, 'xls');
}
