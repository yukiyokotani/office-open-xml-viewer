import { describe, expect, it } from 'vitest';
import { readLegacyDocSourceModuleConfig, readNativePictureBulletPolicy } from './doc-picture-bullet-policy.js';
import { readLegacySourceModuleConfig } from './source-module-config.js';

const wasmUrl = 'https://example.test/direct-doc.wasm';
describe('explicit DOC picture-bullet reading policy', () => {
  it('defaults to strict and transports the closed reading value', () => {
    expect(readNativePictureBulletPolicy(undefined)).toBe('strict');
    expect(readLegacyDocSourceModuleConfig({ wasmUrl, nativePictureBulletPolicy: 'storedSizeForReading' }))
      .toEqual({ wasmUrl, maxInputBytes: 256 * 1024 * 1024, nativeContourPolicy: 'strict',
        nativeWordBreakingPolicy: 'strict', nativePictureBulletPolicy: 'storedSizeForReading' });
  });
  it.each([null, true, 1, '', 'auto', 'storedSize', {}, ['strict']])('rejects an unknown policy %j', value => {
    expect(() => readNativePictureBulletPolicy(value)).toThrow(TypeError);
  });
  it('does not broaden XLS/PPT config or tolerate unrelated DOC fields', () => {
    expect(() => readLegacySourceModuleConfig({ wasmUrl, nativePictureBulletPolicy: 'storedSizeForReading' }, 'legacy XLS'))
      .toThrow('unknown field');
    expect(() => readLegacyDocSourceModuleConfig({ wasmUrl, nativePictureBulletPolicy: 'storedSizeForReading', unexpected: 1 }))
      .toThrow('unknown field');
  });
});
