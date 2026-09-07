import { validateLegacySourceDescriptor, type LegacyDirectSourceDescriptor } from './legacy-source-descriptor.js';

/** First-party native decoder selection for PPT. */
export interface LegacyPptDirectSourceDescriptor extends LegacyDirectSourceDescriptor<'ppt'> {}

export const MAX_LEGACY_PPT_SOURCE_BYTES = 256 * 1024 * 1024;

export function validateLegacyPptSourceDescriptor(value: unknown): LegacyPptDirectSourceDescriptor {
  return validateLegacySourceDescriptor(value, 'ppt');
}
