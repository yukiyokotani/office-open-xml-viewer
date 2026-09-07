import { describe, expect, it } from 'vitest';
import { createLegacyXlsSource } from './direct-xls.js';

describe('createLegacyXlsSource', () => {
  it('returns a frozen validated descriptor without initializing WASM', () => {
    const descriptor = createLegacyXlsSource({ wasmUrl: 'https://example.test/direct-xls.wasm' });
    expect(descriptor).toEqual({
      protocol: 'ooxml-legacy-xls-source/v1',
      builtin: 'xls',
      wasmUrl: 'https://example.test/direct-xls.wasm',
    });
    expect(Object.isFrozen(descriptor)).toBe(true);
  });
});
