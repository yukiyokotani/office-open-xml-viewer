import { describe, expect, it } from 'vitest';
import { legacyDocSource } from './legacy-doc.js';
import { readLegacyDocSourceModuleConfig, readLegacySourceModuleConfig } from './source-module-config.js';
const wasmUrl = 'https://example.test/owned-doc.wasm';
describe('explicit native word-breaking policy boundary', () => {
  it('keeps the default descriptor strict and forwards only the explicit reading choice', () => {
    expect(legacyDocSource({ wasmUrl }).beginLoad().module.config).not.toHaveProperty('nativeWordBreakingPolicy');
    const config = legacyDocSource({ wasmUrl, nativeWordBreakingPolicy: 'simplifyForReading' }).beginLoad().module.config;
    expect(readLegacyDocSourceModuleConfig(config).nativeWordBreakingPolicy).toBe('simplifyForReading');
    expect(readLegacyDocSourceModuleConfig({ wasmUrl }).nativeWordBreakingPolicy).toBe('strict');
    expect(() => readLegacySourceModuleConfig(config, 'legacy XLS')).toThrow('unknown field');
  });
  it('refuses unknown policy names, non-scalar config and unrelated fields', () => {
    for (const policy of [null, false, {}, 'normalize', 'normal']) {
      expect(() => readLegacyDocSourceModuleConfig({ wasmUrl, nativeWordBreakingPolicy: policy })).toThrow(TypeError);
    }
    expect(() => readLegacyDocSourceModuleConfig({ wasmUrl, unrelated: true })).toThrow('unknown field');
  });
});
