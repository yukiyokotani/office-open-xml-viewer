import { validateLegacySourceDescriptor, type LegacyDirectSourceDescriptor } from './legacy-source-descriptor.js';

/** First-party native decoder selection for DOC. */
export interface LegacyDocDirectSourceDescriptor extends LegacyDirectSourceDescriptor<'doc'> {}

export const MAX_LEGACY_DOC_SOURCE_BYTES = 256 * 1024 * 1024;

export function validateLegacyDocSourceDescriptor(value: unknown): LegacyDocDirectSourceDescriptor {
  return validateLegacySourceDescriptor(value, 'doc');
}
