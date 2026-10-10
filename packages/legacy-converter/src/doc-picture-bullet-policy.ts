export { readLegacyDocSourceModuleConfig } from './source-module-config.js';

/** DOC-only opt-in; the reading box is not an Office AUTO sizing claim. */
export type NativePictureBulletPolicy = 'strict' | 'storedSizeForReading';
export function readNativePictureBulletPolicy(value: unknown): NativePictureBulletPolicy {
  if (value === undefined || value === 'strict') return 'strict';
  if (value === 'storedSizeForReading') return value;
  throw new TypeError('legacy DOC nativePictureBulletPolicy is invalid');
}
