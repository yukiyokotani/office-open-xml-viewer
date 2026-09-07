import { describe, expect, it } from 'vitest';
import { createLegacyPptSource } from './direct-ppt.js';

describe('createLegacyPptSource', () => {
  it('provides a validated absolute default asset URL', () => {
    const descriptor = createLegacyPptSource();
    expect(new URL(descriptor.wasmUrl).href).toBe(descriptor.wasmUrl);
    expect(descriptor.wasmUrl.length).toBeGreaterThan(0);
    expect(Object.isFrozen(descriptor)).toBe(true);
  });

  it('returns a validated frozen descriptor without loading the runtime', () => {
    const descriptor = createLegacyPptSource({ wasmUrl: 'https://example.test/direct-ppt.wasm' });
    expect(descriptor).toEqual({
      protocol: 'ooxml-legacy-ppt-source/v1',
      builtin: 'ppt',
      wasmUrl: 'https://example.test/direct-ppt.wasm',
    });
    expect(Object.isFrozen(descriptor)).toBe(true);
  });
});
